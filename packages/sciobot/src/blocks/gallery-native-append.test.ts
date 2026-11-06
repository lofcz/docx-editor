import { expect, test } from 'bun:test';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  galleryDecodePort,
  galleryPng,
  nativeGalleryDocument,
  nativeGalleryDrawings,
} from '../../test/support/native-gallery';
import { openGallery } from '../../test/support/gallery';
import { groupAdjacentImages } from './gallery-group';
import { readGallery, updateGallery } from './gallery';
import { appendNativeGalleryImages } from './gallery-native-append';
import { readGalleryImageAsset } from './gallery-native-asset';
import { removeBlock } from './remove';
import * as Y from 'yjs';

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`native gallery upload preserves duplicate resources and independent peer ownership ${clientIds}`, async () => {
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
      const append = (editor: typeof alice, caption: string) =>
        appendNativeGalleryImages(editor!, {
          blockId: grouped.blockId,
          images: [
            { reference: 'https://assets.example.test/same.png', caption, alt: caption + ' alt' },
          ],
          resolve: async () => ({ bytes: galleryPng, mime: 'image/png' }),
          decodePort: galleryDecodePort,
        });
      pair.pause();
      expect(await append(alice, 'Alice\ncaption')).toMatchObject({ success: true });
      expect(await append(bob, 'Bob caption')).toMatchObject({ success: true });
      pair.reconnect();
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage(),
          gallery = readGallery(pkg, grouped.blockId)!;
        expect(gallery.images).toHaveLength(5);
        expect(new Set(gallery.images.map((image) => image.id)).size).toBe(5);
        expect(gallery.images.map((image) => image.attributes.caption)).toContain('Alice\ncaption');
        expect(gallery.images.map((image) => image.attributes.caption)).toContain('Bob caption');
        for (const image of gallery.images) {
          const asset = readGalleryImageAsset(pkg, grouped.blockId, image.id)!;
          expect(asset.bytes).toEqual(galleryPng);
          expect(asset.path).toMatch(/^media\/[a-f0-9]{64}\.png$/);
          expect(image.attributes).toEqual(asset.attributes);
        }
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(
        readGallery(bob!.surface!.session.currentPackage(), grouped.blockId)!.images.map(
          (image) => image.attributes.caption
        )
      ).toEqual(['First image', 'Second image', 'Third image', 'Bob caption']);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = openGallery(new Uint8Array(await editor.save()));
        expect(readGallery(pkg, grouped.blockId)!.images).toHaveLength(5);
        for (const image of readGallery(pkg, grouped.blockId)!.images)
          expect(readGalleryImageAsset(pkg, grouped.blockId, image.id)!.bytes).toEqual(galleryPng);
      }
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
      const added = readGallery(
        alice!.surface!.session.currentPackage(),
        grouped.blockId
      )!.images.slice(3);
      for (const image of added) {
        expect(
          updateGallery(alice!, {
            blockId: grouped.blockId,
            change: { type: 'remove', imageId: image.id },
          }).success
        ).toBe(true);
        await pair.settle();
      }
      for (const editor of pair.editors) {
        const reopened = openGallery(new Uint8Array(await editor.save()));
        expect(readGallery(reopened, grouped.blockId)!.images).toHaveLength(3);
        expect(
          [...reopened.partBytes.keys()].filter((name) => name.includes('/media/'))
        ).toHaveLength(3);
      }
    } finally {
      pair.destroy();
    }
  });

for (const action of ['abort', 'delete'] as const)
  test(`native gallery upload refuses ${action} during preparation without resource changes`, async () => {
    const pair = await createTwoEditorFixture({
      document: await nativeGalleryDocument(),
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
      let release!: () => void;
      const waiting = new Promise<void>((resolve) => (release = resolve)),
        controller = new AbortController();
      const pending = appendNativeGalleryImages(alice!, {
        blockId: grouped.blockId,
        images: [{ reference: 'file-chosen-by-user' }],
        signal: controller.signal,
        resolve: async () => {
          await waiting;
          return { bytes: galleryPng, mime: 'image/png' };
        },
        decodePort: galleryDecodePort,
      });
      if (action === 'abort') controller.abort();
      else {
        expect(
          removeBlock(bob!, {
            controlId: readGallery(bob!.surface!.session.currentPackage(), grouped.blockId)!
              .controlId,
            mode: 'delete',
          }).success
        ).toBe(true);
        await pair.settle();
      }
      const before = alice!.surface!.session.currentPackage();
      const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      release();
      expect((await pending).success).toBe(false);
      expect(alice!.surface!.session.currentPackage()).toBe(before);
      expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
      for (const editor of pair.editors) {
        const reopened = openGallery(new Uint8Array(await editor.save()));
        expect([...reopened.partBytes].filter(([name]) => name.includes('/media/'))).toEqual(
          [...before.partBytes].filter(([name]) => name.includes('/media/'))
        );
        if (action === 'delete') expect(readGallery(reopened, grouped.blockId)).toBeNull();
      }
    } finally {
      pair.destroy();
    }
  });

test('a pending native gallery upload refuses changed content without inserting resources', async () => {
  const pair = await createTwoEditorFixture({
    document: await nativeGalleryDocument(),
    ownedAttributes: true,
    editor: { imageDecodePort: galleryDecodePort },
  });
  try {
    const [alice, bob] = pair.editors;
    const grouped = groupAdjacentImages(alice!, {
      drawingNodeId: nativeGalleryDrawings(alice!.surface!.session.currentPackage())[0]!.drawing.id,
    });
    if (!grouped.success) throw Error(grouped.error);
    await pair.settle();
    let release!: () => void;
    const waiting = new Promise<void>((resolve) => (release = resolve));
    const pending = appendNativeGalleryImages(alice!, {
      blockId: grouped.blockId,
      images: [{ reference: 'file-chosen-by-user' }],
      resolve: async () => {
        await waiting;
        return { bytes: galleryPng, mime: 'image/png' };
      },
      decodePort: galleryDecodePort,
    });
    const image = readGallery(bob!.surface!.session.currentPackage(), grouped.blockId)!.images[0]!;
    expect(
      updateGallery(bob!, {
        blockId: grouped.blockId,
        change: {
          type: 'caption',
          imageId: image.id,
          expectedCaption: 'First image',
          caption: 'Peer changed',
        },
      }).success
    ).toBe(true);
    await pair.settle();
    const before = alice!.surface!.session.currentPackage();
    release();
    expect(await pending).toMatchObject({ success: false, code: 'stale-revision' });
    expect(alice!.surface!.session.currentPackage()).toBe(before);
    expect(readGallery(before, grouped.blockId)!.images).toHaveLength(3);
  } finally {
    pair.destroy();
  }
});
