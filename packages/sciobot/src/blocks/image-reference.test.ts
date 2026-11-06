import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode, type OoxmlNode } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { zipDocument } from '../../../pro/src/collaboration/__tests__/document-peer-support';
import { galleryPng, galleryDecodePort } from '../../test/support/native-gallery';
import { retainNativeImageReference, type NativeImageReference } from './image-reference';

function findDrawing(editor: DocxEditorInstance): OoxmlNode | null {
  const pending: OoxmlNode[] = [editor.surface!.session.part().root];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'drawing') return node;
    if (node.kind !== 'textValue') for (const child of node.children) pending.push(child);
  }
  return null;
}
function drawing(editor: DocxEditorInstance): OoxmlNode {
  const found = findDrawing(editor);
  if (!found) throw Error('The expected drawing is absent.');
  return found;
}
const bytes = zipDocument('<w:p/><w:p><w:r><w:t>Peer text</w:t></w:r></w:p>');
const insert = (editor: DocxEditorInstance) =>
  editor.executeImageCommand({
    type: 'insertImage',
    data: galleryPng,
    mime: 'image/png',
    widthPoints: 36,
    heightPoints: 36,
  });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    test(`retained image references follow exact publication and never follow identical replacements ${clientIds}/${first}`, async () => {
      const pair = await createTwoEditorFixture({
        document: bytes,
        clientIds,
        editor: { imageDecodePort: galleryDecodePort },
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        pair.pause();
        let source: string | undefined;
        let captured: NativeImageReference | undefined;
        const capture = alice.surface!.session.subscribe(() => {
          const node = findDrawing(alice);
          if (captured || !node) return;
          source = node.id;
          captured = retainNativeImageReference(alice, source);
        });
        expect((await insert(alice)).ok).toBe(true);
        capture();
        if (!captured || !source) throw Error('The local native insertion was not observed.');
        const reference = captured;
        expect(alice.exec({ type: 'setAlignment', align: 'right' }).ok).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        const shared = drawing(alice).id;
        expect(shared).not.toBe(source);
        expect(findNode(alice.surface!.session.part(), source)).toBeNull();
        expect(reference.read()).toBe(shared);
        for (let round = 0; round < 3; round++) {
          const paragraphId = bob.surface!.session.paragraphIds()[1]!;
          expect(
            bob.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId, offset: 0, text: String(round) },
            ]).committed
          ).toBe(true);
          await pair.settle();
          expect(reference.read()).toBe(drawing(alice).id);
        }
        const disposable = retainNativeImageReference(alice, drawing(alice).id);
        disposable.dispose();
        expect(disposable.read()).toBeNull();
        expect(
          bob.surface!.selectDrawing(drawing(bob).id, bob.surface!.session.paragraphIds()[0]!)
        ).toBe(true);
        pair.pause();
        expect(bob.exec({ type: 'deleteImage' }).ok).toBe(true);
        expect((await insert(bob)).ok).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        expect(drawing(alice).id).not.toBe(shared);
        expect(reference.read()).toBeNull();
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(reference.read()).toBeNull();
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(reference.read()).toBeNull();
        const replaced = retainNativeImageReference(alice, drawing(alice).id);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        alice.load(await alice.save());
        expect(replaced.read()).toBeNull();
      } finally {
        pair.destroy();
      }
    });
