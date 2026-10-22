import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { nodeChildren } from '../../test/support/native-table-oracle';
import { protectedDocument } from '../../test/support/protected-document';
import {
  targetSourceFixture,
  projectSourcePackage,
  sourceRevision,
  sourceDefinitions,
  sourceParagraph,
  sourceControl,
} from '../../test/support/target-source-fixture';
import { projectMdx } from '../projection/mdx';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { planSiblingBlocks } from './sibling-blocks';
import {
  indexTargetSource,
  TARGET_SOURCE_LIMITS,
  type TargetSourceInput,
  type TargetSourceIndex,
} from './target-source';

const fixture = targetSourceFixture();
function indexed(input: TargetSourceInput): TargetSourceIndex {
  const result = indexTargetSource(input);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.index;
}
const part = (input: TargetSourceInput) => input.pkg.parts.get(input.pkg.mainDocumentPart)!;
const mutable = (input: TargetSourceInput): TargetSourceInput => ({
  ...input,
  revision: { ...input.revision },
  projection: JSON.parse(JSON.stringify(input.projection)),
});
const ownerText = (input: TargetSourceInput, index: TargetSourceIndex, text: string) =>
  index.owners.find(
    (owner) => owner.kind === 'paragraph' && paragraphTextOf(part(input), owner.id) === text
  )!;
const ownerBlock = (index: TargetSourceIndex, id: string) =>
  index.owners.find((owner) => owner.block?.id === id)!;
const native = (input: TargetSourceInput, id: string) => {
  const node = findNode(part(input), id);
  if (!node) throw new Error(`Missing native node ${id}`);
  return node;
};
function frozen(value: unknown): boolean {
  return (
    value === null ||
    typeof value !== 'object' ||
    (Object.isFrozen(value) && Object.values(value).every(frozen))
  );
}
function parsed(text: string): TargetAstNode[] {
  const result = parseTargetMarkdown(text);
  if (!result.success) throw new Error(result.error);
  return [...result.target.normalizedAst.children];
}

test('owners keep native immediate parents, original child positions, complete markup extents, and physical source slices', async () => {
  const input = await fixture,
    index = indexed(input),
    before = ownerText(input, index, 'Before');
  expect(index.bodyId).toBe(bodyStoryRoot(part(input))!.id);
  expect(index.partName).toBe(input.pkg.mainDocumentPart);
  expect(before).toMatchObject({
    kind: 'paragraph',
    parentId: index.bodyId,
    childIndex: 0,
    paragraphId: before.id,
    cell: null,
    visibility: 'projected',
    block: null,
    opaque: null,
    checklist: null,
  });
  expect(before.extents).toEqual([{ start: 0, end: 10, native: { start: 0, end: 6 } }]);
  expect(input.projection.text.slice(before.extents[0]!.start, before.extents[0]!.end)).toBe(
    '**Before**'
  );
  expect(before.sources).toEqual(
    input.projection.sourceMap.sources.filter(
      (source) => source.sourceScope === 'body' && source.paragraphId === before.id
    )
  );
  for (const owner of index.owners)
    expect(nodeChildren(native(input, owner.parentId))[owner.childIndex]!.id).toBe(owner.id);
  expect(frozen(index)).toBe(true);
});

test('sibling windows retain every native gap, label, property holder, and section anchor', async () => {
  const input = await fixture,
    index = indexed(input),
    body = native(input, index.bodyId),
    all = nodeChildren(body),
    root = index.windows.filter((window) => window.parentId === index.bodyId);
  expect(root.map((window) => ({ start: window.startIndex, count: window.ids.length }))).toEqual([
    { start: 0, count: 1 },
    { start: 2, count: 7 },
    { start: 10, count: 1 },
  ]);
  expect(root[0]!.beforeId).toBeNull();
  expect(root[0]!.afterId).toBe(all[1]!.id);
  expect(root[1]!.beforeId).toBe(all[1]!.id);
  expect(root[1]!.afterId).toBe(all[9]!.id);
  expect(root[2]!.afterId).toBe(all[11]!.id);
  for (const window of index.windows) {
    const children = nodeChildren(native(input, window.parentId));
    expect(window.ids).toEqual(
      children
        .slice(window.startIndex, window.startIndex + window.ids.length)
        .map((node) => node.id)
    );
    expect(window.beforeId).toBe(children[window.startIndex - 1]?.id ?? null);
    expect(window.afterId).toBe(children[window.startIndex + window.ids.length]?.id ?? null);
    for (const id of window.ids)
      expect(index.owners.find((owner) => owner.id === id)?.visibility).toBe('projected');
  }
  for (const owner of index.owners.filter(
    (owner) => owner.visibility === 'generated' || owner.visibility === 'atomic-fallback'
  ))
    expect(index.windows.some((window) => window.ids.includes(owner.id))).toBe(false);
});

test('nested domain content starts after its generated label and retains separate canonical ownership', async () => {
  const input = await fixture,
    index = indexed(input),
    phase = ownerBlock(index, 'phase'),
    activity = ownerBlock(index, 'activity'),
    question = ownerText(input, index, 'Question');
  const phaseContent = contentControlContentNodeOf(native(input, phase.id))!,
    activityContent = contentControlContentNodeOf(native(input, activity.id))!;
  expect(phase.block?.kind).toBe('container');
  expect(activity.parentId).toBe(phaseContent.id);
  expect(question.parentId).toBe(activityContent.id);
  expect(index.windows.find((window) => window.parentId === phaseContent.id)).toMatchObject({
    startIndex: 1,
    ids: [activity.id],
    beforeId: nodeChildren(phaseContent)[0]!.id,
    afterId: null,
  });
  const window = index.windows.find((window) => window.parentId === activityContent.id)!;
  expect(window.startIndex).toBe(1);
  expect(window.ids).toHaveLength(4);
  expect(window.ids[0]).toBe(question.id);
  expect(ownerText(input, index, 'Phase label').visibility).toBe('generated');
  expect(ownerText(input, index, 'Activity label').visibility).toBe('generated');
});

test('complete opaque ASTs retain native block ownership without exposing their fallback or inline run containers', async () => {
  const input = await fixture,
    index = indexed(input),
    opaque = ownerBlock(index, 'opaque'),
    inline = ownerBlock(index, 'inline'),
    stored = buildBlockIndex(input.pkg);
  expect<unknown>(opaque.opaque?.ast).toEqual(stored.byId.get('opaque')!.payload.mdast);
  expect(opaque.opaque?.ast.data).toEqual({
    future: { type: 'payload', position: { keep: true } },
  });
  expect(opaque.block?.kind).toBe('atomic');
  expect(opaque.paragraphId).toBeNull();
  expect(ownerText(input, index, 'Atomic fallback').visibility).toBe('atomic-fallback');
  expect(inline.opaque?.inline).toBe(true);
  expect(inline.opaque?.ast.future).toEqual({ keep: true });
  expect(inline.paragraphId).toBe(inline.parentId);
  expect(index.windows.some((window) => window.ids.includes(inline.id))).toBe(false);
  const content = contentControlContentNodeOf(native(input, inline.id))!;
  expect(index.windows.some((window) => window.parentId === content.id)).toBe(false);
  expect((globalThis as Record<string, unknown>).__targetSourceExecuted).toBeUndefined();
});

test('custom registry kinds distinguish an empty container from an atomic block', async () => {
  const input = await fixture,
    index = indexed(input),
    container = ownerBlock(index, 'custom_container'),
    atom = ownerBlock(index, 'custom_atomic');
  expect(container.block?.kind).toBe('container');
  expect(atom.block?.kind).toBe('atomic');
  expect(container.block!.contentStart).toBe(container.block!.contentEnd);
  expect(atom.block!.contentStart).toBe(atom.block!.contentEnd);
  const content = contentControlContentNodeOf(native(input, container.id))!,
    fallback = contentControlContentNodeOf(native(input, atom.id))!;
  const empty = index.owners.find((owner) => owner.parentId === content.id)!;
  expect(empty.visibility).toBe('projected');
  expect(empty.extents[0]!.start).toBe(empty.extents[0]!.end);
  expect(index.windows.find((window) => window.parentId === content.id)?.ids).toEqual([empty.id]);
  expect(index.windows.some((window) => window.parentId === fallback.id)).toBe(false);
});

test('a host registry override keeps atomic ownership even for a built-in container type', async () => {
  const input = await fixture,
    definitions = new Map(sourceDefinitions);
  definitions.set('lesson_activity', { ...definitions.get('lesson_activity')!, kind: 'atomic' });
  const source = {
      ...input,
      projection: await projectSourcePackage(input.pkg, 'proposed', definitions),
    },
    index = indexed(source),
    activity = ownerBlock(index, 'activity'),
    content = contentControlContentNodeOf(native(source, activity.id))!;
  expect(activity.block?.kind).toBe('atomic');
  expect(ownerText(source, index, 'Question').visibility).toBe('atomic-fallback');
  expect(index.windows.some((window) => window.parentId === content.id)).toBe(false);
  expect(index.owners.some((owner) => owner.block?.id === 'opaque')).toBe(false);
});

test('merged and nested tables keep physical scopes without fabricated text for hidden continuations', async () => {
  const input = await fixture,
    index = indexed(input),
    outer = index.owners.find((owner) => owner.kind === 'table' && owner.cell === null)!,
    inner = index.owners.find((owner) => owner.kind === 'table' && owner.cell !== null)!,
    outerText = ownerText(input, index, 'Outer cell'),
    nestedText = ownerText(input, index, 'Nested A');
  expect(inner.parentId).toBe(outerText.parentId);
  expect(inner.cell).toEqual(outerText.cell);
  expect(inner.cell?.tableId).toBe(outer.id);
  expect(nestedText.cell?.tableId).toBe(inner.id);
  expect(nestedText.parentId).toBe(nestedText.cell!.cellId);
  expect(nestedText.cell?.cellId).not.toBe(outerText.cell?.cellId);
  const continuation = index.owners.find(
    (owner) =>
      owner.kind === 'paragraph' &&
      owner.cell?.tableId === outer.id &&
      owner.visibility === 'unprojected'
  )!;
  expect(continuation.extents).toEqual([]);
  expect(continuation.sources).toEqual([]);
  expect(index.windows.some((window) => window.ids.includes(continuation.id))).toBe(false);
  expect(index.windows.find((window) => window.parentId === outerText.parentId)?.ids).toEqual([
    outerText.id,
    inner.id,
  ]);
});

test('empty task metadata retains the generated placeholder and exact native checkbox range', async () => {
  const input = await fixture,
    index = indexed(input),
    owner = index.owners.find((owner) => owner.checklist)!,
    task = owner.checklist!;
  expect(task.empty).toBe(true);
  expect(task.checked).toBe(false);
  expect(task.native).toEqual({ start: 0, end: 1 });
  expect(input.projection.text.slice(task.start, task.end)).toContain('\u00a0');
  expect(owner.paragraphId).toBe(task.paragraphId);
  expect(owner.id).toBe(task.paragraphId);
  expect(owner.extents).toEqual([
    { start: task.start, end: task.end, native: { start: 0, end: 1 } },
  ]);
  const content = contentControlContentNodeOf(native(input, task.controlId))!;
  expect(index.windows.some((window) => window.parentId === content.id)).toBe(false);
});

test('ordinary Word controls keep their native wrapper while inner paragraphs receive their own window', async () => {
  const input = await fixture,
    index = indexed(input),
    owner = ownerText(input, index, 'Word form text'),
    content = native(input, owner.parentId);
  expect(content.kind).toBe('contentControlContent');
  expect(owner.visibility).toBe('projected');
  expect(index.windows.find((window) => window.parentId === owner.parentId)?.ids).toEqual([
    owner.id,
  ]);
  expect(
    index.windows
      .filter((window) => window.parentId === index.bodyId)
      .some((window) => window.ids.includes(owner.id))
  ).toBe(false);
});

test('existing sibling planner consumes a returned window with original anchors and no whole-body replacement', async () => {
  const input = await fixture,
    index = indexed(input),
    owner = ownerText(input, index, 'Tail'),
    window = index.windows.find((window) => window.ids.includes(owner.id))!,
    extent = owner.extents[0]!,
    ast = parsed(input.projection.text.slice(extent.start, extent.end))[0]!;
  const result = planSiblingBlocks(
    {
      revision: input.revision,
      part: part(input),
      parentId: window.parentId,
      startIndex: window.startIndex,
      blocks: [{ id: owner.id, ast }],
    },
    [...parsed('New paragraph'), ast]
  );
  if (!result.success) throw new Error(result.error);
  expect(result.plan.operations).toEqual([
    {
      kind: 'siblings',
      partName: part(input).name,
      parentId: window.parentId,
      beforeId: window.beforeId,
      afterId: owner.id,
      removedIds: [],
      inserted: parsed('New paragraph'),
    },
  ]);
  expect(result.pairs).toMatchObject([{ id: owner.id, changed: false }]);
});

test('metadata-only empty bodies and generated-label-only containers retain zero-length insertion anchors', async () => {
  const loaded = readOoxmlPackage(
    docxFixture('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>')
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const input = {
      revision: sourceRevision,
      pkg: loaded.package,
      projection: await projectSourcePackage(loaded.package),
    },
    index = indexed(input),
    body = bodyStoryRoot(part(input))!;
  expect(index.owners).toEqual([]);
  expect(index.windows).toEqual([
    {
      parentId: body.id,
      startIndex: 0,
      ids: [],
      beforeId: null,
      afterId: nodeChildren(body)[0]!.id,
    },
  ]);
  const insertion = planSiblingBlocks(
    { revision: input.revision, part: part(input), parentId: body.id, startIndex: 0, blocks: [] },
    parsed('New body.')
  );
  expect(insertion.success).toBe(true);
  const doc = readOoxmlPackage(
    docxFixture(
      sourceControl(
        'empty',
        `<w:sdt><w:sdtPr><w:tag w:val="scio:label"/></w:sdtPr><w:sdtContent>${sourceParagraph('Label')}</w:sdtContent></w:sdt>`
      )
    )
  );
  if (!doc.ok) throw new Error(doc.reason);
  const stored = withPayloadRecord(
    doc.package,
    {
      id: 'empty',
      type: 'lesson_activity',
      version: 1,
      attributes: { name: 'Empty', duration: 0 },
    },
    ''
  );
  if (!stored.ok) throw new Error(stored.reason);
  const empty = {
      revision: sourceRevision,
      pkg: stored.pkg,
      projection: await projectSourcePackage(stored.pkg),
    },
    mapped = indexed(empty),
    control = ownerBlock(mapped, 'empty'),
    content = contentControlContentNodeOf(native(empty, control.id))!;
  expect(mapped.windows.find((window) => window.parentId === content.id)).toEqual({
    parentId: content.id,
    startIndex: 1,
    ids: [],
    beforeId: nodeChildren(content)[0]!.id,
    afterId: null,
  });
});

test.each(['original', 'proposed', 'all-markup'] as const)(
  '%s source indexing preserves imported headers, footers, notes, media, and sections',
  async (mode) => {
    const loaded = readOoxmlPackage(protectedDocument());
    if (!loaded.ok) throw new Error(loaded.reason);
    const pkg = loaded.package,
      input = { revision: sourceRevision, pkg, projection: await projectSourcePackage(pkg, mode) },
      before = writeOoxmlPackage(pkg),
      fingerprints = [...pkg.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)]);
    const index = indexed(input);
    expect(index.owners.every((owner) => owner.id.startsWith(pkg.mainDocumentPart + '#'))).toBe(
      true
    );
    expect(writeOoxmlPackage(pkg)).toEqual(before);
    expect([...pkg.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)])).toEqual(
      fingerprints
    );
    expect(
      input.projection.sourceMap.paragraphs.some((source) => source.sourceScope !== 'body')
    ).toBe(true);
    const fallbacks = input.projection.sourceMap.sources.filter(
      (source) => !source.exact && source.sourceScope === 'body'
    );
    expect(fallbacks.length).toBeGreaterThan(0);
    for (const fallback of fallbacks) {
      expect(
        index.owners.find((owner) => owner.id === fallback.paragraphId)?.sources
      ).toContainEqual(fallback);
      if (fallback.markdownBoundaries === undefined)
        expect(Object.hasOwn(fallback, 'markdownBoundaries')).toBe(false);
    }
  }
);

test('stale, malformed, and conflicting metadata refuses without a partial source index', async () => {
  const input = await fixture,
    before = writeOoxmlPackage(input.pkg);
  const changes: ((value: TargetSourceInput) => void)[] = [
    (s) => {
      (s.projection.revision as { revision: number }).revision++;
    },
    (s) => {
      (s.projection as unknown as Record<string, unknown>).blocks = [
        s.projection.blocks[0],
        s.projection.blocks[0],
      ];
    },
    (s) => {
      (s.projection.blocks[0] as { controlId: string }).controlId = 'foreign';
    },
    (s) => {
      (s.projection.blocks[0] as { contentEnd: number }).contentEnd = s.projection.text.length + 1;
    },
    (s) => {
      (s.projection.sourceMap.paragraphs[0] as { paragraphId: string }).paragraphId = 'foreign';
    },
    (s) => {
      (s.projection.sourceMap.paragraphs[0] as { sourceEnd: number }).sourceEnd = 999;
    },
    (s) => {
      (s.projection.sourceMap.sources[0] as { sourceEnd: number }).sourceEnd = 999;
    },
    (s) => {
      (s.projection.sourceMap.sources[0] as unknown as Record<string, unknown>).markdownBoundaries =
        [{ sourceOffset: 0, markdownOffset: -1 }];
    },
    (s) => {
      (s.projection.sourceMap.tables[0] as { tableId: string }).tableId =
        s.projection.sourceMap.paragraphs[0]!.paragraphId;
    },
    (s) => {
      (s.projection.opaqueMdx[0] as { inline: boolean }).inline = true;
    },
    (s) => {
      (s.projection.opaqueMdx[0]!.ast as unknown as Record<string, unknown>).future = 'forged';
    },
    (s) => {
      (s.projection as unknown as Record<string, unknown>).fallbackParagraphIds = [
        s.projection.sourceMap.paragraphs[0]!.paragraphId,
      ];
    },
    (s) => {
      (s.projection.checklists[0] as { controlId: string }).controlId =
        s.projection.blocks[0]!.controlId;
    },
    (s) => {
      (s.projection.checklists[0]!.native as { end: number }).end = 0;
    },
    (s) => {
      (s.projection as unknown as Record<string, unknown>).checklists = [
        s.projection.checklists[0],
        s.projection.checklists[0],
      ];
    },
  ];
  for (const change of changes) {
    const value = mutable(input);
    change(value);
    const result = indexTargetSource(value);
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'index')).toBe(false);
  }
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
});

test('getters, cycles, depth/node/byte excess, and source descriptor traps refuse before evaluation', async () => {
  const input = await fixture;
  let calls = 0;
  for (const field of ['revision', 'pkg', 'projection']) {
    const source = Object.defineProperty({ ...input }, field, {
      enumerable: true,
      get() {
        calls++;
        return input[field as keyof TargetSourceInput];
      },
    });
    expect(indexTargetSource(source).success).toBe(false);
  }
  const invalid: unknown[] = [];
  const getter = mutable(input);
  Object.defineProperty(getter.projection, 'text', {
    enumerable: true,
    get() {
      calls++;
      return '';
    },
  });
  invalid.push(getter.projection);
  const cycle: Record<string, unknown> = {};
  cycle.next = cycle;
  invalid.push(cycle);
  const deep: Record<string, unknown> = { ...mutable(input).projection };
  let node = deep;
  for (let i = 0; i < TARGET_SOURCE_LIMITS.maxDepth; i++) {
    const next = {};
    node.extra = next;
    node = next;
  }
  invalid.push(deep);
  invalid.push({
    ...mutable(input).projection,
    blocks: new Array(TARGET_SOURCE_LIMITS.maxNodes + 1),
  });
  invalid.push({ ...mutable(input).projection, text: 'x'.repeat(TARGET_SOURCE_LIMITS.maxBytes) });
  for (const projection of invalid)
    expect(
      indexTargetSource({ ...input, projection: projection as TargetSourceInput['projection'] })
        .success
    ).toBe(false);
  expect(calls).toBe(0);
});

test('duplicate/cyclic canonical identities and excessive native depth refuse before index construction', async () => {
  const input = await fixture,
    original = part(input),
    body = bodyStoryRoot(original)!;
  const altered = (replacement: OoxmlNode): OoxmlPart => ({
    ...original,
    root: {
      ...original.root,
      children: nodeChildren(original.root).map((node) =>
        node.id === body.id ? replacement : node
      ),
    } as OoxmlPart['root'],
  });
  const first = nodeChildren(body)[0]!,
    duplicate = { ...body, children: [first, first] } as OoxmlNode;
  const cyclic = { ...body, children: [] } as Exclude<OoxmlNode, { kind: 'textValue' }>;
  (cyclic as unknown as { children: OoxmlNode[] }).children = [cyclic];
  let deep = first;
  for (let i = 0; i < TARGET_SOURCE_LIMITS.maxCanonicalDepth + 1; i++)
    deep = {
      id: `deep-${i}`,
      kind: 'generic',
      namespaceUri: 'urn:keep',
      localName: 'wrapper',
      attributes: [],
      namespaceBindings: [],
      children: [deep],
    } as OoxmlNode;
  for (const changed of [duplicate, cyclic, { ...body, children: [deep] } as OoxmlNode]) {
    const pkg = withPart(input.pkg, altered(changed)),
      result = indexTargetSource({ ...input, pkg });
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'index')).toBe(false);
  }
});

test('outputs detach without freezing caller metadata and source reads preserve live history and package bytes', async () => {
  const value = mutable(await fixture),
    before = JSON.stringify(value.projection),
    index = indexed(value);
  expect(JSON.stringify(value.projection)).toBe(before);
  expect(Object.isFrozen(value.projection)).toBe(false);
  expect(frozen(index)).toBe(true);
  (value.projection.blocks[0] as { tag: string }).tag = 'changed';
  expect(
    index.owners.find((owner) => owner.block?.id === value.projection.blocks[0]!.id)?.block?.tag
  ).not.toBe('changed');
  const original = await fixture,
    editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(original.pkg),
    });
  try {
    const surface = editor.surface!,
      session = surface.session;
    surface.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: session.paragraphIds()[0]!, offset: 0, text: 'User ' },
    ]);
    const revision = { ...sourceRevision, revision: session.packageRevision() },
      opened = openDocumentForExport(session, { displayMode: 'proposed' });
    if (!opened.ok) throw new Error(opened.reason);
    try {
      const projected = projectMdx(
        session.currentPackage(),
        await opened.session.layout(),
        revision,
        sourceDefinitions
      );
      if (!projected.success) throw new Error(projected.error);
      const source = { revision, pkg: session.currentPackage(), projection: projected.projection },
        bytes = writeOoxmlPackage(source.pkg),
        selection = surface.state().selection,
        history = [session.canUndo(), session.canRedo()];
      expect(indexTargetSource(source).success).toBe(true);
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
      expect(session.packageRevision()).toBe(revision.revision);
      expect(surface.state().selection).toEqual(selection);
      expect([session.canUndo(), session.canRedo()]).toEqual(history);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(paragraphTextOf(session.part(), session.paragraphIds()[0]!)).toBe('Before');
    } finally {
      opened.session.dispose();
    }
  } finally {
    editor.destroy();
  }
});
