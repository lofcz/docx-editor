import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { protectedDocument } from '../../test/support/protected-document';
import { docxFixture } from '../../test/support/document';
import { directParagraphTextOps, planParagraphText } from '../diff/paragraph-text';
import { collectProtectedContent, checkProtectedContent } from './protected-content';
import { projectMarkdown } from './markdown';
import { resolvePatchSpan } from './resolve-patch';

async function imported(bytes = protectedDocument(), mode: 'original' | 'proposed' = 'proposed') {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
  });
  const session = editor.surface!.session;
  const opened = openDocumentForExport(session);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const revision = {
      documentId: 'protected',
      epoch: 'load-1',
      revision: session.packageRevision(),
    };
    const projection = projectMarkdown(await opened.session.layoutFor(mode), revision);
    const result = collectProtectedContent(session.currentPackage(), projection, revision);
    if (!result.success) throw new Error(result.error);
    return {
      editor,
      session,
      revision,
      projection,
      index: result.index,
      resolve(text: string) {
        const start = projection.text.indexOf(text);
        if (start < 0) throw new Error(`Missing text: ${text}`);
        return resolvePatchSpan(
          projection,
          { start, end: start + text.length },
          { revision, protected: result.index.boundaries }
        );
      },
    };
  } catch (error) {
    editor.destroy();
    throw error;
  } finally {
    opened.session.dispose();
  }
}
function fingerprints(pkg: OoxmlPackage) {
  return [...pkg.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]);
}
/** Corrupt an immutable preview in tests, without touching the live package. */
function changedNode(
  pkg: OoxmlPackage,
  id: string,
  change: (node: OoxmlNode) => OoxmlNode
): OoxmlPackage {
  const visit = (node: OoxmlNode): OoxmlNode => {
    if (node.id === id) return change(node);
    if (node.kind === 'textValue') return node;
    const children = node.children.map(visit);
    return children.some((child, at) => child !== node.children[at])
      ? ({ ...node, children } as OoxmlNode)
      : node;
  };
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  return withPart(pkg, { ...part, root: visit(part.root) } as OoxmlPart);
}
function appendForeignAttribute(node: OoxmlNode): OoxmlNode {
  if (node.kind === 'textValue') throw new Error('Expected element');
  return {
    ...node,
    attributes: [
      ...node.attributes,
      {
        kind: 'genericExtension',
        namespaceUri: 'urn:opaque',
        localName: 'changed',
        prefix: 'x',
        value: 'bad',
      },
    ],
  } as OoxmlNode;
}

test('the pure inventory retains exact protected references and detached immutable data', async () => {
  const value = await imported();
  try {
    const { session, index, projection, revision } = value;
    const before = fingerprints(session.currentPackage());
    const bytes = writeOoxmlPackage(session.currentPackage());
    const selection = value.editor.surface!.state().selection;
    expect(projection.text).toContain('Before');
    expect(projection.text).toContain('Cached');
    expect(projection.text).toContain('Left right');
    expect(projection.text).not.toContain('opaque block');
    expect(new Set(index.nodes.map((ref) => ref.kind))).toEqual(
      new Set([
        'element-shell',
        'opaque-xml',
        'field',
        'drawing',
        'inline-atom',
        'section-properties',
        'outside-body',
      ])
    );
    expect(index.parts.map((ref) => ref.partName)).toEqual(
      expect.arrayContaining([
        '/word/header1.xml',
        '/word/footer1.xml',
        '/word/footnotes.xml',
        '/custom/opaque.xml',
        '/custom/opaque.bin',
        '/word/media/image.png',
      ])
    );
    expect(index.nodes.filter((ref) => ref.kind === 'section-properties')).toHaveLength(2);
    expect(index.boundaries.filter((ref) => ref.reason === 'Word field')).toHaveLength(2);
    expect(index.boundaries.filter((ref) => ref.reason === 'Word drawing')).toHaveLength(1);
    expect(index.boundaries.some((ref) => ref.sourceScope.startsWith('footnote'))).toBe(true);
    expect(value.resolve('Protected note')).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    expect(
      Object.isFrozen(index) && Object.isFrozen(index.nodes) && Object.isFrozen(index.boundaries)
    ).toBe(true);
    expect(index.nodes.every(Object.isFrozen) && index.parts.every(Object.isFrozen)).toBe(true);
    expect(JSON.parse(JSON.stringify(index))).toEqual(index);
    expect(collectProtectedContent(session.currentPackage(), projection, revision)).toEqual({
      success: true,
      index,
    });
    expect(checkProtectedContent(index, session.currentPackage(), revision)).toEqual({
      success: true,
    });
    expect(fingerprints(session.currentPackage())).toEqual(before);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(revision.revision);
    expect(value.editor.surface!.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual([false, false]);
  } finally {
    value.editor.destroy();
  }
});

test('native body text edits beside fields and opaque inline XML preserve protected content through save/reopen', async () => {
  for (const [before, after] of [
    ['Before', 'Changed'],
    ['after', 'later'],
    ['Left ', 'Ahead '],
    ['right', 'opposite'],
    ['tail', 'ending'],
    ['Photo', 'Picture'],
  ]) {
    const value = await imported();
    try {
      const { session, revision, index } = value;
      const resolved = value.resolve(before!);
      if (!resolved.success) throw new Error(resolved.error);
      if (resolved.target.kind !== 'text')
        throw new Error(`Expected text, got ${resolved.target.kind}`);
      const target = resolved.target;
      const result = planParagraphText(
        {
          revision,
          partName: session.part().name,
          paragraphId: target.start.paragraphId,
          text: before!,
          offset: target.start.offset,
        },
        after!
      );
      if (!result.success) throw new Error(result.error);
      const original = session.currentPackage();
      const originalFingerprint = fingerprints(original);
      let candidate = original;
      const ops = result.plan.operations.flatMap((edit) =>
        edit.kind === 'text' ? directParagraphTextOps(edit) : []
      );
      for (const op of ops) {
        const applied = applyTreeOp(candidate.parts.get(candidate.mainDocumentPart)!, op);
        if (!applied.ok) throw new Error(applied.reason);
        candidate = withPart(candidate, applied.part);
      }
      expect(checkProtectedContent(index, candidate, revision)).toEqual({ success: true });
      expect(fingerprints(original)).toEqual(originalFingerprint);
      expect(value.editor.surface!.applyAutomationOps(() => ops).rejected).toBe(false);
      expect(checkProtectedContent(index, session.currentPackage(), revision)).toEqual({
        success: true,
      });
      expect(paragraphTextOf(session.part(), target.start.paragraphId)).toContain(after!);
      const reloaded = readOoxmlPackage(writeOoxmlPackage(session.currentPackage()));
      if (!reloaded.ok) throw new Error(reloaded.reason);
      expect(fingerprints(reloaded.package)).toEqual(fingerprints(session.currentPackage()));
      for (const [name, bytes] of original.partBytes)
        if (!original.parts.has(name) && name !== original.mainDocumentPart)
          expect(reloaded.package.partBytes.get(name)).toEqual(bytes);
      expect(session.canUndo()).toBe(true);
      session.undo();
      expect(fingerprints(session.currentPackage())).toEqual(originalFingerprint);
      session.redo();
      expect(paragraphTextOf(session.part(), target.start.paragraphId)).toContain(after!);
    } finally {
      value.editor.destroy();
    }
  }
});

test('field caches, omitted drawings, and crossings identify their canonical protected boundary', async () => {
  const value = await imported();
  try {
    for (const text of ['7', 'Cached', 'Left right']) {
      const result = value.resolve(text);
      expect(result.success).toBe(false);
      if (result.success) throw new Error('Expected refusal');
      expect(result.code).toBe('unsupported-operation');
      expect(result.error).toContain('protected content (');
    }
    const drawing = value.index.boundaries.find((ref) => ref.reason === 'Word drawing')!;
    // The default tool projection omits images. Its native atom still owns this boundary.
    expect(drawing.end).toBe(drawing.start);
    expect(value.resolve('Photo  end')).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    const opaqueBlock = value.index.boundaries.find(
      (ref) => !ref.native && ref.sourceScope === 'body'
    )!;
    const crossBlock = resolvePatchSpan(
      value.projection,
      {
        start: value.projection.text.indexOf('right'),
        end: value.projection.text.indexOf('Complex') + 7,
      },
      { revision: value.revision, protected: [opaqueBlock] }
    );
    expect(crossBlock).toMatchObject({ success: false, code: 'unsupported-operation' });
    if (!crossBlock.success) expect(crossBlock.error).toContain(opaqueBlock.nodeId!);
    expect(
      resolvePatchSpan(
        value.projection,
        { start: drawing.start, end: drawing.end },
        { revision: value.revision, protected: value.index.boundaries }
      )
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    for (const ref of value.index.boundaries.filter((ref) => ref.start === ref.end)) {
      const result = resolvePatchSpan(
        value.projection,
        { start: ref.start, end: ref.end },
        { revision: value.revision, protected: [ref] }
      );
      expect(result).toMatchObject({ success: false, code: 'unsupported-operation' });
      if (!result.success) expect(result.error).toContain(ref.nodeId!);
    }
  } finally {
    value.editor.destroy();
  }
});

test('immutable previews refuse each changed protected subtree or shell before any live write', async () => {
  const value = await imported();
  try {
    const original = value.session.currentPackage(),
      before = fingerprints(original);
    for (const reference of value.index.nodes) {
      const candidate = changedNode(original, reference.nodeId, appendForeignAttribute);
      const result = checkProtectedContent(value.index, candidate, value.revision);
      expect(result).toMatchObject({
        success: false,
        code: 'unsupported-operation',
        protectedReference: { nodeId: reference.nodeId },
      });
    }
    expect(fingerprints(original)).toEqual(before);
    expect(value.session.packageRevision()).toBe(value.revision.revision);
    expect(value.session.canUndo()).toBe(false);
  } finally {
    value.editor.destroy();
  }
});

test('untouched XML, binary parts, and OPC metadata remain protected', async () => {
  const value = await imported();
  try {
    const original = value.session.currentPackage(),
      before = writeOoxmlPackage(original);
    for (const ref of value.index.parts) {
      let candidate: OoxmlPackage;
      if (ref.mode === 'xml') {
        const part = original.parts.get(ref.partName)!;
        expect(
          checkProtectedContent(
            value.index,
            withPart(original, {
              ...part,
              root: { ...part.root, namespaceBindings: [] },
            }),
            value.revision
          )
        ).toMatchObject({ success: false, protectedReference: { partName: ref.partName } });
        candidate = withPart(original, {
          ...part,
          root: appendForeignAttribute(part.root),
        } as OoxmlPart);
      } else {
        const partBytes = new Map(original.partBytes);
        const bytes = partBytes.get(ref.partName)!.slice();
        bytes[0] = bytes[0]! ^ 1;
        partBytes.set(ref.partName, bytes);
        candidate = { ...original, partBytes };
      }
      expect(checkProtectedContent(value.index, candidate, value.revision)).toMatchObject({
        success: false,
        code: 'unsupported-operation',
        protectedReference: { partName: ref.partName },
      });
    }
    const relationships = new Map(original.relationships);
    relationships.delete(original.mainDocumentPart);
    expect(
      checkProtectedContent(value.index, { ...original, relationships }, value.revision)
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    const defaults = new Map(original.contentTypes.defaults);
    defaults.set('fake', 'application/fake');
    expect(
      checkProtectedContent(
        value.index,
        { ...original, contentTypes: { ...original.contentTypes, defaults } },
        value.revision
      )
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(writeOoxmlPackage(original)).toEqual(before);
    expect(value.session.canUndo()).toBe(false);
  } finally {
    value.editor.destroy();
  }
});

test('protection rejects stale snapshots and invalid projection addresses', async () => {
  const value = await imported();
  try {
    const current = { ...value.revision, revision: value.revision.revision + 1 };
    expect(
      collectProtectedContent(value.session.currentPackage(), value.projection, current)
    ).toMatchObject({ success: false, code: 'stale-revision' });
    expect(
      checkProtectedContent(value.index, value.session.currentPackage(), current)
    ).toMatchObject({ success: false, code: 'stale-revision' });
    const source = value.projection.sourceMap.sources[0]!;
    const projection = {
      ...value.projection,
      sourceMap: { ...value.projection.sourceMap, sources: [{ ...source, markdownEnd: Infinity }] },
    };
    expect(
      collectProtectedContent(value.session.currentPackage(), projection, value.revision)
    ).toMatchObject({ success: false, code: 'invalid-document' });
    const paragraph = value.projection.sourceMap.paragraphs[0]!;
    expect(
      collectProtectedContent(
        value.session.currentPackage(),
        {
          ...value.projection,
          sourceMap: {
            ...value.projection.sourceMap,
            paragraphs: [{ ...paragraph, markdownEnd: NaN }],
          },
        },
        value.revision
      )
    ).toMatchObject({ success: false, code: 'invalid-document' });
  } finally {
    value.editor.destroy();
  }
});

test('protected parent ownership, sibling order, and namespace bindings survive body patches', async () => {
  const value = await imported();
  try {
    const original = value.session.currentPackage(),
      part = value.session.part();
    const ref = value.index.nodes.find(
      (ref) =>
        ref.reason === 'Unprojected OOXML' &&
        findNode(part, ref.nodeId)?.kind === 'generic' &&
        (findNode(part, ref.nodeId) as { localName: string }).localName === 'opaque'
    )!;
    expect(Boolean(ref)).toBe(true);
    const parent = findNode(part, ref.parentId!)!;
    if (parent.kind === 'textValue') throw new Error('Missing parent');
    const removed = changedNode(original, parent.id, (node) =>
      node.kind === 'textValue'
        ? node
        : ({
            ...node,
            children: node.children.filter((child) => child.id !== ref.nodeId),
          } as OoxmlNode)
    );
    expect(checkProtectedContent(value.index, removed, value.revision)).toMatchObject({
      success: false,
      protectedReference: { nodeId: ref.nodeId },
    });
    const moved = changedNode(original, parent.id, (node) =>
      node.kind === 'textValue'
        ? node
        : ({ ...node, children: [...node.children].reverse() } as OoxmlNode)
    );
    const orderResult = checkProtectedContent(value.index, moved, value.revision);
    expect(orderResult).toMatchObject({ success: false, code: 'unsupported-operation' });
    if (!orderResult.success)
      expect(orderResult.error).toContain('reorders protected parent content');
    const destination = value.session.paragraphIds().find((id) => id !== parent.id)!;
    const reparented = changedNode(removed, destination, (node) =>
      node.kind === 'textValue'
        ? node
        : ({ ...node, children: [...node.children, findNode(part, ref.nodeId)!] } as OoxmlNode)
    );
    const parentResult = checkProtectedContent(value.index, reparented, value.revision);
    expect(parentResult).toMatchObject({
      success: false,
      protectedReference: { nodeId: ref.nodeId },
    });
    if (!parentResult.success) expect(parentResult.error).toContain('Protected node parent');
    for (const reference of value.index.nodes.filter((ref) => {
      const node = findNode(part, ref.nodeId);
      return node?.kind !== 'textValue' && node?.namespaceBindings.length;
    })) {
      const candidate = changedNode(original, reference.nodeId, (node) =>
        node.kind === 'textValue' ? node : ({ ...node, namespaceBindings: [] } as OoxmlNode)
      );
      expect(checkProtectedContent(value.index, candidate, value.revision)).toMatchObject({
        success: false,
        protectedReference: { nodeId: reference.nodeId },
      });
    }
    expect(value.session.canUndo()).toBe(false);
  } finally {
    value.editor.destroy();
  }
});

test.each(['original', 'proposed'] as const)(
  'hidden fields keep zero-width protection in the %s view',
  async (mode) => {
    const tag = mode === 'original' ? 'ins' : 'del';
    const bytes = docxFixture(
      '<w:p><w:r><w:t>Before</w:t></w:r>' +
        `<w:${tag} w:id="9" w:author="Ada"><w:fldSimple w:instr="PAGE"><w:r><w:t>7</w:t></w:r></w:fldSimple></w:${tag}>` +
        '<w:r><w:t>After</w:t></w:r></w:p>'
    );
    const value = await imported(bytes, mode);
    try {
      expect(value.projection.text).toBe('BeforeAfter');
      const field = value.index.boundaries.find((ref) => ref.reason === 'Word field')!;
      expect(field).toMatchObject({ start: 6, end: 6, edges: 'interior' });
      expect(value.resolve('Before').success).toBe(true);
      expect(value.resolve('After').success).toBe(true);
      expect(value.resolve('BeforeAfter')).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
      expect(
        checkProtectedContent(value.index, value.session.currentPackage(), value.revision)
      ).toEqual({ success: true });
    } finally {
      value.editor.destroy();
    }
  }
);

test('binary fingerprints detect in-place changes without retaining mutable bytes in the index', async () => {
  const value = await imported();
  try {
    const pkg = value.session.currentPackage();
    const name = '/custom/opaque.bin',
      bytes = pkg.partBytes.get(name)!;
    const previous = bytes[0]!;
    try {
      bytes[0] = previous ^ 1;
      expect(checkProtectedContent(value.index, pkg, value.revision)).toMatchObject({
        success: false,
        code: 'unsupported-operation',
        protectedReference: { partName: name },
      });
    } finally {
      bytes[0] = previous;
    }
    expect(checkProtectedContent(value.index, pkg, value.revision)).toEqual({ success: true });
  } finally {
    value.editor.destroy();
  }
});
