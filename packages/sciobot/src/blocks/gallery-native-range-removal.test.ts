import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { openGallery } from '../../test/support/gallery';
import {
  galleryDecodePort,
  nativeGalleryDocument,
  nativeGalleryDrawings,
} from '../../test/support/native-gallery';
import { groupAdjacentImages } from './gallery-group';
import { readGallery, updateGallery } from './gallery';
import { removeBlockSelection } from './remove';

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const includeTail of [false, true])
    test(`native gallery range deletion preserves peer edits and owned resources (${clientIds}, tail=${includeTail})`, async () => {
      const pair = await createTwoEditorFixture({
        document: await nativeGalleryDocument(),
        clientIds,
        ownedAttributes: true,
        editor: { imageDecodePort: galleryDecodePort },
      });
      try {
        const [alice, bob] = pair.editors;
        const grouped = groupAdjacentImages(alice!, {
          drawingNodeId: nativeGalleryDrawings(alice!.surface!.session.currentPackage())[0]!.drawing
            .id,
        });
        if (!grouped.success) throw Error(grouped.error);
        await pair.settle();
        const gallery = readGallery(alice!.surface!.session.currentPackage(), grouped.blockId)!;
        const tail = alice!.surface!.session.paragraphIds().at(-1)!;
        const resources = [...alice!.surface!.session.currentPackage().partBytes].filter(([name]) =>
          name.includes('/media/')
        );
        const drawings = nativeGalleryDrawings(alice!.surface!.session.currentPackage()).map(
          ({ drawing }) => canonicalOoxmlFingerprint(drawing)
        );
        pair.pause();
        expect(
          updateGallery(bob!, {
            blockId: gallery.id,
            change: {
              type: 'caption',
              imageId: gallery.images[0]!.id,
              expectedCaption: 'First image',
              caption: 'Caption during removal',
            },
          }).success
        ).toBe(true);
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: tail, offset: 4, text: ' peer' },
          ]).committed
        ).toBe(true);
        const peerOnly = semanticDigest(bob!.surface!.session.currentPackage().parts.values());
        const result = removeBlockSelection(alice!, {
          nodeIds: includeTail ? [gallery.controlId, tail] : [gallery.controlId],
        });
        if (!result.success) throw Error(result.error);
        expect(result.removedPayloadIds).toHaveLength(7);
        pair.reconnect(clientIds[0] === 41 ? 0 : 1);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(readGallery(editor.surface!.session.currentPackage(), gallery.id)).toBeNull();
          expect(nativeGalleryDrawings(editor.surface!.session.currentPackage())).toEqual([]);
          expect(Boolean(findNode(editor.surface!.session.part(), tail))).toBe(!includeTail);
          expect(editor.surface!.session.paragraphIds().length).toBeGreaterThan(0);
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          const pkg = editor.surface!.session.currentPackage();
          expect(semanticDigest(pkg.parts.values())).toEqual(peerOnly);
          expect(readGallery(pkg, gallery.id)!.images[0]!.attributes.caption).toBe(
            'Caption during removal'
          );
          expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Tail peer');
          expect(
            nativeGalleryDrawings(pkg).map(({ drawing }) => canonicalOoxmlFingerprint(drawing))
          ).toEqual(drawings);
          expect([...pkg.partBytes].filter(([name]) => name.includes('/media/'))).toEqual(
            resources
          );
          const reopened = openGallery(new Uint8Array(await editor.save()));
          expect(readGallery(reopened, gallery.id)!.images[0]!.attributes.caption).toBe(
            'Caption during removal'
          );
        }
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        for (const editor of pair.editors)
          expect(
            readGallery(openGallery(new Uint8Array(await editor.save())), gallery.id)
          ).toBeNull();
      } finally {
        pair.destroy();
      }
    });
