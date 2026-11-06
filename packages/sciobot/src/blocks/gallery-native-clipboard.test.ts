import { expect, test } from 'bun:test';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { openGallery } from '../../test/support/gallery';
import {
  galleryDecodePort,
  galleryDescendants,
  galleryPng,
  nativeGalleryDocument,
  nativeGalleryDrawings,
} from '../../test/support/native-gallery';
import { collectBlockClipboardData, pasteWithBlocks } from '../react-blocks/clipboard';
import { groupAdjacentImages } from './gallery-group';
import { readGallery, updateGallery } from './gallery';
import { readGalleryImageAsset } from './gallery-native-asset';
import { buildBlockIndex } from './block-index';
import { removeBlock } from './remove';

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`native gallery clipboard retains distinct drawing and caption owners ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await nativeGalleryDocument(),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: galleryDecodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const result = groupAdjacentImages(alice!, {
        drawingNodeId: nativeGalleryDrawings(alice!.surface!.session.currentPackage())[0]!.drawing
          .id,
      });
      if (!result.success) throw Error(result.error);
      await pair.settle();
      const transfer = (editor: typeof alice) => {
        const surface = editor!.surface!,
          ids = surface.session.paragraphIds();
        const end = ids[5]!;
        surface.setSelection({
          anchor: { paragraphId: ids[0]!, offset: 0 },
          head: { paragraphId: end, offset: paragraphTextOf(surface.session.part(), end)!.length },
        });
        const data = collectBlockClipboardData(editor!)!;
        expect(data.blocks).toHaveLength(7);
        const flavours = surface.copyFlavours();
        const tail = { paragraphId: ids[6]!, offset: 4 };
        surface.setSelection({ anchor: tail, head: tail });
        return { ...flavours, data };
      };
      const a = transfer(alice),
        b = transfer(bob);
      pair.pause();
      expect(pasteWithBlocks(alice!, a, { createId: (payload) => 'a-' + payload!.id })).toBe(true);
      expect(pasteWithBlocks(bob!, b, { createId: (payload) => 'b-' + payload!.id })).toBe(true);
      pair.reconnect();
      await pair.settle();
      const assertCopies = (pkg: ReturnType<typeof openGallery>) => {
        expect(buildBlockIndex(pkg).diagnostics).toEqual([]);
        const drawings = nativeGalleryDrawings(pkg);
        expect(drawings).toHaveLength(9);
        expect(new Set(drawings.map(({ drawing }) => drawing.id)).size).toBe(9);
        const nativeIds = drawings.flatMap(({ drawing }) =>
          galleryDescendants(drawing)
            .filter((node) => node.kind !== 'textValue' && node.localName === 'docPr')
            .map(
              (node) =>
                node.kind !== 'textValue' &&
                node.attributes.find((attribute) => attribute.localName === 'id')!.value
            )
        );
        expect(new Set(nativeIds).size).toBe(9);
        for (const prefix of ['', 'a-', 'b-']) {
          const gallery = readGallery(pkg, prefix + result.blockId)!;
          expect(gallery.images).toHaveLength(3);
          expect(gallery.images.every((image) => !prefix || image.id.startsWith(prefix))).toBe(
            true
          );
          for (const image of gallery.images)
            expect(readGalleryImageAsset(pkg, gallery.id, image.id)!.bytes).toEqual(galleryPng);
        }
      };
      for (const editor of pair.editors) assertCopies(editor.surface!.session.currentPackage());
      const copied = readGallery(alice!.surface!.session.currentPackage(), 'a-' + result.blockId)!;
      expect(
        updateGallery(alice!, {
          blockId: copied.id,
          change: {
            type: 'caption',
            imageId: copied.images[0]!.id,
            expectedCaption: 'First image',
            caption: result.blockId,
          },
        }).success
      ).toBe(true);
      await pair.settle();
      expect(
        readGallery(bob!.surface!.session.currentPackage(), result.blockId)!.images[0]!.attributes
          .caption
      ).toBe('First image');
      expect(
        removeBlock(bob!, {
          controlId: readGallery(bob!.surface!.session.currentPackage(), 'b-' + result.blockId)!
            .controlId,
          mode: 'delete',
        }).success
      ).toBe(true);
      await pair.settle();
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const reopened = openGallery(new Uint8Array(await editor.save()));
        assertCopies(reopened);
        expect(readGallery(reopened, copied.id)!.images[0]!.attributes.caption).toBe(
          result.blockId
        );
      }
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(
          readGallery(openGallery(new Uint8Array(await editor.save())), 'b-' + result.blockId)
        ).toBeNull();
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
    } finally {
      pair.destroy();
    }
  });
