import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode, type OoxmlNode } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { zipDocument } from '../../../pro/src/collaboration/__tests__/document-peer-support';

const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (character) => character.charCodeAt(0)
);
const decodePort = { decode: async () => ({ pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 }) };
const bytes = zipDocument('<w:p/><w:p><w:r><w:t>Peer text</w:t></w:r></w:p>');
function drawing(editor: DocxEditorInstance): OoxmlNode {
  const visit = (node: OoxmlNode): OoxmlNode | undefined => {
    if (node.kind === 'drawing') return node;
    if (node.kind !== 'textValue')
      for (const child of node.children) {
        const found = visit(child);
        if (found) return found;
      }
  };
  const result = visit(editor.surface!.session.part().root);
  if (!result) throw Error('Expected an image.');
  return result;
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`image selection follows shared identity through alignment and reconnect (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({
      document: bytes,
      clientIds,
      editor: { imageDecodePort: decodePort },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      expect(
        (
          await alice.executeImageCommand({
            type: 'insertImage',
            data: png,
            mime: 'image/png',
            widthPoints: 36,
            heightPoints: 36,
          })
        ).ok
      ).toBe(true);
      pair.flush();
      const source = drawing(alice).id;
      const paragraphId = alice.surface!.session.paragraphIds()[0]!;
      expect(alice.surface!.selectDrawing(source, paragraphId)).toBe(true);
      expect(alice.snapshot().image?.id).toBe(source);
      pair.pause();
      expect(alice.exec({ type: 'setAlignment', align: 'right' }).ok).toBe(true);
      const peerParagraph = bob.surface!.session.paragraphIds()[1]!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: peerParagraph, offset: 9, text: ' kept' },
        ]).committed
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      const shared = drawing(alice).id;
      expect(shared).not.toBe(source);
      expect(findNode(alice.surface!.session.part(), source)).toBeNull();
      expect(alice.surface!.drawingSelectionIntent()).toEqual({
        kind: 'pointer',
        drawingNodeId: shared,
      });
      expect(alice.snapshot().image?.id).toBe(shared);
      expect(alice.snapshot().formatting?.alignment).toBe('right');
      expect(alice.snapshot().image?.id).toBe(drawing(bob).id);
      const aligned = await pair.savedDigests();
      expect(aligned[0]).toEqual(aligned[1]);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(alice.surface!.session.bodyText()).toContain('Peer text kept');
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(await pair.savedDigests()).toEqual(aligned);
    } finally {
      pair.destroy();
    }
  });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`peer replacement never transfers drawing selection to another image (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({
      document: bytes,
      clientIds,
      editor: { imageDecodePort: decodePort },
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      expect(
        (
          await alice.executeImageCommand({
            type: 'insertImage',
            data: png,
            mime: 'image/png',
            widthPoints: 36,
            heightPoints: 36,
          })
        ).ok
      ).toBe(true);
      await pair.settle();
      expect(
        alice.surface!.selectDrawing(drawing(alice).id, alice.surface!.session.paragraphIds()[0]!)
      ).toBe(true);
      expect(
        bob.surface!.selectDrawing(drawing(bob).id, bob.surface!.session.paragraphIds()[0]!)
      ).toBe(true);
      pair.pause();
      expect(alice.exec({ type: 'setAlignment', align: 'right' }).ok).toBe(true);
      expect(bob.exec({ type: 'deleteImage' }).ok).toBe(true);
      expect(
        (
          await bob.executeImageCommand({
            type: 'insertImage',
            data: png,
            mime: 'image/png',
            widthPoints: 24,
            heightPoints: 24,
          })
        ).ok
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      expect(drawing(alice).id).toBe(drawing(bob).id);
      expect(alice.snapshot().image).toBeNull();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(bob.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(await pair.savedDigests()).toEqual(saved);
      expect(alice.snapshot().image).toBeNull();
    } finally {
      pair.destroy();
    }
  });
