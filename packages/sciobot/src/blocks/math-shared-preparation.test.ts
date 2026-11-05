import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  isAtomicContentControl,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { mathDocument } from '../../test/support/math-document';
import { buildBlockIndex } from './block-index';
import { insertBlock } from './insert';
import { createMathPayload } from './math-codec';
import { insertMathAtSelection, updateMath } from './math-edit';
import { insertMath } from './math-insert';
import { updateAtomicBlock } from './update-atomic';

async function legacyDocument() {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: mathDocument(),
  });
  try {
    expect(
      insertBlock(editor, {
        parentControlId: null,
        beforeId: null,
        payload: createMathPayload('legacy', 'block', 'first\nsecond'),
        label: 'Equation',
        paragraphs: ['first', 'second'],
      }).success
    ).toBe(true);
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
  }
}

for (const selected of [true, false])
  test(`shared math insertion preserves existing unprepared equation ownership: selected=${selected}`, async () => {
    const pair = await createTwoEditorFixture({ document: await legacyDocument() });
    try {
      const editor = pair.editors[0]!;
      const surface = editor.surface!;
      const before = buildBlockIndex(surface.session.currentPackage()).byId.get('legacy')!;
      const legacy = findNode(surface.session.part(), before.controlId)!;
      expect(isAtomicContentControl(legacy)).toBe(false);
      const point = { paragraphId: surface.session.paragraphIds()[0]!, offset: 0 };
      const inserted = selected
        ? insertMathAtSelection(editor, {
            id: 'fresh',
            source: 'new',
            kind: 'inline',
            selection: { anchor: point, head: point },
            expectedRevision: surface.session.packageRevision(),
          })
        : insertMath(editor, {
            payload: createMathPayload('fresh', 'inline', 'new'),
            position: { kind: 'inline', ...point },
          });
      expect(inserted.success).toBe(true);
      await pair.settle();
      for (const peer of pair.editors) {
        const index = buildBlockIndex(peer.surface!.session.currentPackage());
        expect(findNode(peer.surface!.session.part(), index.byId.get('legacy')!.controlId)).toEqual(
          legacy
        );
        expect(
          isAtomicContentControl(
            findNode(peer.surface!.session.part(), index.byId.get('fresh')!.controlId)!
          )
        ).toBe(true);
      }
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const peer of pair.editors) {
        const index = buildBlockIndex(peer.surface!.session.currentPackage());
        expect(index.byId.has('fresh')).toBe(false);
        expect(findNode(peer.surface!.session.part(), index.byId.get('legacy')!.controlId)).toEqual(
          legacy
        );
      }
    } finally {
      pair.destroy();
    }
  });

test('unprepared shared equations refuse source edits and conversion without changing shared history', async () => {
  const pair = await createTwoEditorFixture({ document: await legacyDocument() });
  try {
    const editor = pair.editors[0]!;
    const surface = editor.surface!;
    const before = surface.session.currentPackage();
    const fingerprint = canonicalOoxmlFingerprint(surface.session.part());
    const original = buildBlockIndex(before).byId.get('legacy')!.payload;
    for (const kind of ['inline', 'block'] as const)
      expect(updateMath(editor, { original, source: 'changed', kind })).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
    expect(
      updateAtomicBlock(editor, {
        blockId: original.id,
        attributes: { ...original.attributes, texExpression: 'changed' },
      })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(surface.session.currentPackage()).toBe(before);
    await pair.settle();
    for (const peer of pair.editors)
      expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toBe(fingerprint);
    expect(editor.exec({ type: 'undo' }).ok).toBe(false);
  } finally {
    pair.destroy();
  }
});
