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
import { groupAdjacentImages, readAdjacentGalleryImages } from './gallery-group';
import { readGallery, updateGallery } from './gallery';
import { ungroupGallery } from './gallery-ungroup';
import { readPayloadRecords } from './payload-opc-codec';
import { readNativeGalleryImage } from './gallery-native-record';
import { updateInlineImage } from './inline-image-update';

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`group and ungroup retain native images, captions, peer edits, and history ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await nativeGalleryDocument(),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: galleryDecodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const original = alice!.surface!.session.currentPackage();
      const images = nativeGalleryDrawings(original);
      const adjacent = readAdjacentGalleryImages(
        original,
        original.mainDocumentPart,
        images[1]!.drawing.id
      );
      expect(adjacent).toHaveLength(3);
      const paragraphs = adjacent.flatMap((image) => [
        image.paragraphId,
        ...(image.captionId ? [image.captionId] : []),
      ]);
      const before = paragraphs.map((id) =>
        canonicalOoxmlFingerprint(findNode(alice!.surface!.session.part(), id)!)
      );
      const media = [...original.partBytes].filter(([name]) => name.includes('/media/'));
      pair.pause();
      const grouped = groupAdjacentImages(alice!, { drawingNodeId: images[1]!.drawing.id });
      if (!grouped.success) throw Error(grouped.error);
      const remoteCaption = adjacent[1]!.captionId!;
      const peer = bob!.surface!.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: remoteCaption, offset: 12, text: ' peer' }],
        { kind: 'body' },
        undefined,
        false,
        true
      );
      expect(peer.rejected).toBe(false);
      pair.reconnect();
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage();
        const gallery = readGallery(pkg, grouped.blockId)!;
        expect(gallery.payload.version).toBe(3);
        expect(gallery.images.map((image) => image.attributes.caption)).toEqual([
          'First image',
          'Second image peer',
          'Third image',
        ]);
        expect([...pkg.partBytes].filter(([name]) => name.includes('/media/'))).toEqual(media);
        const records = readPayloadRecords(pkg);
        expect(records.status).toBe('ready');
        if (records.status !== 'ready') throw Error('Missing metadata');
        for (const image of gallery.images) {
          const record = records.entries.get(image.id)!.decoded;
          expect(record.status).toBe('supported');
          if (record.status === 'supported')
            expect(record.value.attributes).toEqual({ resource: { version: 1, kind: 'embedded' } });
        }
        expect(gallery.images[0]!.attributes).toMatchObject({
          width: 64,
          height: 32,
          caption: 'First image',
        });
        const native = readNativeGalleryImage(
          pkg,
          editor.surface!.session.part(),
          findNode(editor.surface!.session.part(), gallery.images[0]!.controlId)!
        )!;
        expect(native.fields.drawing.id).toBe(images[0]!.drawing.id);
        expect(canonicalOoxmlFingerprint(native.paragraph)).toBe(before[0]!);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(readGallery(alice!.surface!.session.currentPackage(), grouped.blockId)).toBeNull();
      expect(paragraphTextOf(bob!.surface!.session.part(), remoteCaption)).toBe(
        'Second image peer'
      );
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const current = nativeGalleryDrawings(bob!.surface!.session.currentPackage())[0]!;
      expect(
        updateInlineImage(bob!, {
          drawingNodeId: current.drawing.id,
          expectedRevision: bob!.surface!.session.packageRevision(),
          change: { type: 'alt', alt: 'Peer alt' },
        }).success
      ).toBe(true);
      await pair.settle();
      const gallery = readGallery(alice!.surface!.session.currentPackage(), grouped.blockId)!;
      expect(gallery.images[0]!.attributes.alt).toBe('Peer alt');
      expect(
        updateGallery(alice!, {
          blockId: grouped.blockId,
          change: {
            type: 'move',
            imageId: gallery.images[2]!.id,
            beforeImageId: gallery.images[0]!.id,
            expectedOrder: gallery.images.map((image) => image.id),
          },
        }).success
      ).toBe(true);
      await pair.settle();
      const ungroup = await ungroupGallery(alice!, {
        blockId: grouped.blockId,
        decodePort: galleryDecodePort,
        resolve: async () => {
          throw Error('Native galleries must not fetch');
        },
      });
      if (!ungroup.success) throw Error(ungroup.error);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = openGallery(new Uint8Array(await editor.save()));
        expect(readGallery(pkg, grouped.blockId)).toBeNull();
        expect(
          nativeGalleryDrawings(pkg).map((image) => image.block.payload.attributes.alt)
        ).toEqual(['Third image', 'Peer alt', 'Second image']);
        expect([...pkg.partBytes].filter(([name]) => name.includes('/media/'))).toEqual(media);
      }
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(
        readGallery(bob!.surface!.session.currentPackage(), grouped.blockId)!.images[0]!.attributes
          .caption
      ).toBe('Third image');
      expect(semanticDigest(alice!.surface!.session.currentPackage().parts.values())).toEqual(
        semanticDigest(bob!.surface!.session.currentPackage().parts.values())
      );
    } finally {
      pair.destroy();
    }
  });
