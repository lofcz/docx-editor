import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { blockClipboardDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { moveBlock } from '../blocks/move';
import { removeBlock } from '../blocks/remove';
import { createStreamAnchor } from './stream-anchor';

const control = (editor: DocxEditorInstance, id: string) =>
  buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(id)!.controlId;
const paragraph = (editor: DocxEditorInstance, text: string) =>
  editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === text)!;

test('caret and selected-sibling anchors write no content, history, or saved placeholders', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: blockClipboardDocument() });
  try {
    const before = new Uint8Array(await editor.save());
    const tree = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const selected = createStreamAnchor(editor, [control(editor, 'activity-1')]);
    if (!selected.success) throw new Error(selected.error);
    expect(selected.anchor.resolve()).toMatchObject({
      success: true,
      point: { afterId: control(editor, 'activity-1') },
    });
    expect(editor.surface!.session.canUndo()).toBe(false);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(tree);
    expect(new Uint8Array(await editor.save())).toEqual(before);
    selected.anchor.release();
    selected.anchor.release();
    expect(selected.anchor.state).toBe('released');
    expect(selected.anchor.resolve().success).toBe(false);

    const id = paragraph(editor, 'Tail');
    expect(id).toBeDefined();
    const intro = paragraph(editor, 'Intro');
    const many = createStreamAnchor(editor, [id, intro]);
    if (!many.success) throw new Error(many.error);
    expect(many.anchor.resolve()).toMatchObject({ success: true, point: { afterId: id } });
    many.anchor.release();
    expect(createStreamAnchor(editor, [id, id]).success).toBe(false);
    expect(createStreamAnchor(editor, [id, control(editor, 'activity-1')]).success).toBe(false);
    expect(new Uint8Array(await editor.save())).toEqual(before);
    editor.surface!.setSelection({
      anchor: { paragraphId: id, offset: 2 },
      head: { paragraphId: id, offset: 2 },
    });
    const caret = createStreamAnchor(editor);
    if (!caret.success) throw new Error(caret.error);
    expect(caret.anchor.resolve()).toMatchObject({
      success: true,
      point: { afterId: id, beforeId: null },
    });
    editor.load(blockClipboardDocument());
    expect(caret.anchor.state).toBe('released');
    expect(caret.anchor.resolve().success).toBe(false);
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const receiver of [0, 1] as const)
  test(`a request anchor follows a peer activity move, then releases after peer deletion; receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({ document: blockClipboardDocument() });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const started = createStreamAnchor(alice, [control(alice, 'activity-1')]);
      if (!started.success) throw new Error(started.error);
      pair.pause();
      expect(
        moveBlock(bob, {
          controlId: control(bob, 'activity-1'),
          parentControlId: control(bob, 'phase-2'),
          beforeId: null,
        }).success
      ).toBe(true);
      const tail = paragraph(alice, 'Tail');
      expect(
        alice.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: tail, offset: 4, text: ' Alice' },
        ]).rejected
      ).toBe(false);
      pair.reconnect(receiver);
      await pair.settle();
      const point = started.anchor.resolve();
      expect(point.success).toBe(true);
      if (!point.success) throw new Error(point.error);
      expect(point.point.afterId).toBe(control(alice, 'activity-1'));
      expect(point.point.beforeId).toBeNull();
      const snapshots = await pair.savedDigests();
      expect(snapshots[0]).toEqual(snapshots[1]);
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(started.anchor.state).toBe('active');
      expect(bob.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(started.anchor.resolve()).toMatchObject({
        success: true,
        point: { afterId: control(alice, 'activity-1') },
      });
      expect(
        removeBlock(bob, { controlId: control(bob, 'activity-1'), mode: 'delete' }).success
      ).toBe(true);
      await pair.settle();
      expect(started.anchor.resolve().success).toBe(false);
      expect(started.anchor.state).toBe('released');
      expect(paragraph(alice, 'Tail Alice')).toBeDefined();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      const reopened = readOoxmlPackage(new Uint8Array(await alice.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(semanticDigest(reopened.package.parts.values())).toEqual(saved[0]);
    } finally {
      pair.destroy();
    }
  });
