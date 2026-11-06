import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  semanticDigest,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { openGallery } from '../../test/support/gallery';
import {
  galleryDescendants,
  galleryDecodePort,
  nativeGalleryDocument,
  nativeGalleryDrawings,
} from '../../test/support/native-gallery';
import { buildBlockIndex } from './block-index';
import { duplicateBlock, duplicateBlockSelection } from './duplicate';
import { groupAdjacentImages } from './gallery-group';
import { readGallery, updateGallery } from './gallery';
import { removeBlock } from './remove';

function pictureIds(pkg: OoxmlPackage) {
  const all = [...pkg.parts.values()].flatMap((part) => galleryDescendants(part.root));
  return {
    nodes: all.filter((node) => node.kind === 'drawing').map((node) => node.id),
    word: all
      .filter((node) => node.kind === 'drawingDocPr')
      .map((node) => node.attributes.find((attribute) => attribute.localName === 'id')!.value),
  };
}
function media(pkg: OoxmlPackage) {
  return [...pkg.partBytes]
    .filter(([name]) => name.includes('/media/'))
    .map(([name, bytes]) => [name, Array.from(bytes)]);
}
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`native gallery copies keep independent drawings, captions, history, and resources (${clientIds})`, async () => {
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
      const source = readGallery(alice!.surface!.session.currentPackage(), grouped.blockId)!;
      const sourceFingerprint = canonicalOoxmlFingerprint(
        findNode(alice!.surface!.session.part(), source.controlId)!
      );
      const sourceMedia = media(alice!.surface!.session.currentPackage());
      pair.pause();
      let a = 0,
        b = 0;
      const first = duplicateBlock(alice!, {
        controlId: source.controlId,
        allocateBlockId: (type) => `alice-${type}-${++a}`,
      });
      const second = duplicateBlockSelection(bob!, {
        nodeIds: [source.controlId],
        allocateBlockId: (type) => `bob-${type}-${++b}`,
      });
      if (!first.success) throw Error(first.error);
      if (!second.success) throw Error(second.error);
      const firstId = first.payloadIds[grouped.blockId]!,
        secondId = second.payloadIds[grouped.blockId]!;
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage(),
          ids = pictureIds(pkg);
        expect(ids.nodes).toHaveLength(9);
        expect(new Set(ids.nodes).size).toBe(9);
        expect(new Set(ids.word).size).toBe(9);
        expect(media(pkg)).toEqual(sourceMedia);
        expect(
          canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), source.controlId)!)
        ).toBe(sourceFingerprint);
        expect(buildBlockIndex(pkg).diagnostics).toEqual([]);
        for (const id of [firstId, secondId]) {
          const copy = readGallery(pkg, id)!;
          expect(copy.images).toHaveLength(3);
          expect(
            copy.images.every(
              (image) => buildBlockIndex(pkg).byId.get(image.id)?.payload.galleryId === id
            )
          ).toBe(true);
          expect(copy.images.map((image) => image.id)).not.toEqual(
            source.images.map((image) => image.id)
          );
        }
      }
      const firstImage = readGallery(alice!.surface!.session.currentPackage(), firstId)!.images[0]!;
      pair.pause();
      expect(
        updateGallery(bob!, {
          blockId: firstId,
          change: {
            type: 'caption',
            imageId: firstImage.id,
            expectedCaption: String(firstImage.attributes.caption),
            caption: 'Peer caption',
          },
        }).success
      ).toBe(true);
      expect(
        removeBlock(alice!, {
          controlId: readGallery(alice!.surface!.session.currentPackage(), firstId)!.controlId,
          mode: 'delete',
        }).success
      ).toBe(true);
      pair.reconnect(clientIds[0] === 41 ? 1 : 0);
      await pair.settle();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(
        readGallery(alice!.surface!.session.currentPackage(), firstId)!.images[0]!.attributes
          .caption
      ).toBe('Peer caption');
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage(),
          reopened = openGallery(new Uint8Array(await editor.save()));
        expect(semanticDigest(reopened.parts.values())).toEqual(semanticDigest(pkg.parts.values()));
        expect(readGallery(reopened, firstId)).toBeNull();
        expect(readGallery(reopened, secondId)!.images).toHaveLength(3);
        const ids = pictureIds(reopened);
        expect(ids.word).toHaveLength(6);
        expect(new Set(ids.word).size).toBe(6);
        expect(media(reopened)).toEqual(sourceMedia);
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
