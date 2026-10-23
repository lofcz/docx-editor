import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { importedMath, mathDocument } from '../../test/support/math-document';
import { buildBlockIndex } from './block-index';
import { createMathPayload, mathWordFallback } from './math-codec';
import { insertMath } from './math-insert';
import { readPayloadStore } from './payload-opc-codec';
const inline = createMathPayload('inline-math', 'inline', '  \\frac{𝛼}{β} < 1  ');
const block = createMathPayload(
  'block-math',
  'block',
  '\\begin{aligned}\r\n x &= y \\\\\n z &= 𝛼\r\\end{aligned}\n'
);
const paragraphs = (node: OoxmlNode): OoxmlNode[] =>
  node.kind === 'textValue'
    ? []
    : node.kind === 'paragraph'
      ? [node]
      : node.children.flatMap(paragraphs);

test('native equations persist exact sources and readable fallbacks; adjacent edits preserve imported OMML through history and DOCX', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: mathDocument(),
  });
  try {
    const surface = editor.surface!;
    const before = canonicalOoxmlFingerprint(surface.session.part());
    const expectedOmml = importedMath(surface.session.currentPackage());
    expect(expectedOmml).toHaveLength(2);
    const first = surface.session.paragraphIds()[0]!;
    const inserted = insertMath(editor, {
      payload: inline,
      position: { kind: 'inline', paragraphId: first, offset: 6 },
    });
    if (!inserted.success) throw Error(inserted.error);
    const inlineTree = canonicalOoxmlFingerprint(surface.session.part());
    const added = insertMath(editor, {
      payload: block,
      position: { kind: 'block', parentControlId: null, beforeId: null },
    });
    if (!added.success) throw Error(added.error);
    const bothTree = canonicalOoxmlFingerprint(surface.session.part());
    const check = async () => {
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw Error(reopened.reason);
      for (const pkg of [surface.session.currentPackage(), reopened.package]) {
        expect(importedMath(pkg)).toEqual(expectedOmml);
        const index = buildBlockIndex(pkg);
        expect(index.diagnostics).toEqual([]);
        for (const payload of [inline, block]) {
          const found = index.byId.get(payload.id)!;
          expect(found.payload).toEqual(payload);
          const part = pkg.parts.get(found.partName)!;
          const node = findNode(part, found.controlId)!;
          if (payload === block)
            expect(paragraphs(node).map((p) => paragraphTextOf(part, p.id))).toEqual(
              mathWordFallback(block).paragraphs
            );
        }
        const part = pkg.parts.get(pkg.mainDocumentPart)!;
        const paragraph = paragraphs(part.root)[0]!;
        expect(paragraphTextOf(part, paragraph.id)).toContain(
          'Before' + mathWordFallback(inline).inlineText + '\uFFFC'
        );
      }
    };
    await check();
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(inlineTree);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
    const store = readPayloadStore(surface.session.currentPackage());
    expect(store.status === 'ready' ? store.entries.size : 0).toBe(0);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(bothTree);
    // Edit each side of the imported equation, without reconstructing the paragraph.
    for (const [offset, text] of [
      [0, 'Lead '],
      [paragraphTextOf(surface.session.part(), first)!.length + 5, ' tail'],
    ] as const) {
      const result = surface.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: first, offset, text }],
        { kind: 'body' },
        [],
        false,
        true
      );
      expect(result.rejected).toBe(false);
      expect(importedMath(surface.session.currentPackage())).toEqual(expectedOmml);
    }
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw Error(saved.reason);
    expect(importedMath(saved.package)).toEqual(expectedOmml);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(bothTree);
    await check();
  } finally {
    editor.destroy();
  }
});

test('invalid, stale, duplicate, and tracked insertion refuse without package writes', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: mathDocument(),
  });
  try {
    const surface = editor.surface!;
    const paragraphId = surface.session.paragraphIds()[0]!;
    const position = { kind: 'inline' as const, paragraphId, offset: 0 };
    for (const request of [
      { payload: block, position },
      { payload: inline, position: { ...position, offset: -1 } },
      { payload: inline, position, expectedRevision: -1 },
    ]) {
      const before = surface.session.currentPackage();
      expect(insertMath(editor, request).success).toBe(false);
      expect(surface.session.currentPackage()).toBe(before);
    }
    expect(insertMath(editor, { payload: inline, position }).success).toBe(true);
    const before = surface.session.currentPackage();
    expect(insertMath(editor, { payload: inline, position }).success).toBe(false);
    expect(surface.session.currentPackage()).toBe(before);
    surface.setEditingMode('suggest');
    expect(
      insertMath(editor, { payload: createMathPayload('tracked', 'inline', 'x'), position })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(surface.session.currentPackage()).toBe(before);
  } finally {
    editor.destroy();
  }
});
