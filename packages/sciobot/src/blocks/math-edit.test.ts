import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  parentNodeOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { mathDocument, importedMath } from '../../test/support/math-document';
import { buildBlockIndex } from './block-index';
import { insertMathAtSelection, updateMath } from './math-edit';

test('selected text becomes math and changes display in one undo unit while retaining equation identity', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: mathDocument(),
  });
  try {
    const surface = editor.surface!;
    const first = surface.session.paragraphIds()[0]!;
    const originalOmml = importedMath(surface.session.currentPackage());
    const before = canonicalOoxmlFingerprint(surface.session.part());
    const insert = insertMathAtSelection(editor, {
      id: 'equation',
      kind: 'inline',
      source: 'x+1',
      selection: {
        anchor: { paragraphId: first, offset: 0 },
        head: { paragraphId: first, offset: 6 },
      },
      expectedRevision: surface.session.packageRevision(),
    });
    expect(insert).toMatchObject({ success: true });
    expect(paragraphTextOf(surface.session.part(), first)).toBe('x+1\uFFFCAfter');
    const block = () => buildBlockIndex(surface.session.currentPackage()).byId.get('equation')!;
    const inline = canonicalOoxmlFingerprint(surface.session.part());
    expect(
      updateMath(editor, { original: block().payload, kind: 'block', source: 'x+2' })
    ).toMatchObject({ success: true });
    expect(parentNodeOf(surface.session.part(), block().controlId)?.kind).toBe('body');
    const display = canonicalOoxmlFingerprint(surface.session.part());
    expect(importedMath(surface.session.currentPackage())).toEqual(originalOmml);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(inline);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(display);
    expect(
      updateMath(editor, { original: block().payload, kind: 'inline', source: '\\sqrt{x}' })
    ).toMatchObject({ success: true });
    expect(parentNodeOf(surface.session.part(), block().controlId)?.kind).toBe('paragraph');
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw Error(saved.reason);
    const read = buildBlockIndex(saved.package);
    expect(read.diagnostics).toEqual([]);
    expect(read.byId.get('equation')!.payload).toEqual(block().payload);
    expect(importedMath(saved.package)).toEqual(originalOmml);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(display);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
  } finally {
    editor.destroy();
  }
});

test('stale math drafts and stale insertion positions refuse without changing document or history', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: mathDocument(),
  });
  try {
    const surface = editor.surface!;
    const paragraphId = surface.session.paragraphIds()[0]!;
    const selection = { anchor: { paragraphId, offset: 3 }, head: { paragraphId, offset: 3 } };
    const request = {
      id: 'equation',
      kind: 'inline' as const,
      source: '',
      selection,
      expectedRevision: surface.session.packageRevision(),
    };
    expect(insertMathAtSelection(editor, request).success).toBe(true);
    const original = buildBlockIndex(surface.session.currentPackage()).byId.get(
      'equation'
    )!.payload;
    expect(updateMath(editor, { original, kind: 'inline', source: 'peer' }).success).toBe(true);
    const before = surface.session.currentPackage();
    expect(updateMath(editor, { original, kind: 'block', source: 'draft' })).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    expect(insertMathAtSelection(editor, request)).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    expect(surface.session.currentPackage()).toBe(before);
    const live = buildBlockIndex(before).byId.get('equation')!;
    expect(findNode(surface.session.part(), live.controlId)).not.toBeNull();
  } finally {
    editor.destroy();
  }
});
