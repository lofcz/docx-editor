import { expect, test } from 'bun:test';
import { findNode, paragraphTextOf, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  galleryDecodePort,
  nativeGalleryDocument,
  nativeGalleryDrawings,
  withoutNativeGalleryCaptions,
} from '../../test/support/native-gallery';
import { openGallery } from '../../test/support/gallery';
import { groupAdjacentImages } from './gallery-group';
import { readGallery, updateGallery } from './gallery';
import { readNativeGalleryImage } from './gallery-native-record';
import { readPayloadRecords } from './payload-opc-codec';
import { updateInlineImage } from './inline-image-update';

const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes].filter(([name]) => name.includes('/media/'));
function nativeCaption(pkg: OoxmlPackage, id: string, at = 0) {
  const gallery = readGallery(pkg, id)!;
  const item = gallery.images[at]!;
  const part = pkg.parts.get(gallery.partName)!;
  const native = readNativeGalleryImage(pkg, part, findNode(part, item.controlId)!)!;
  expect(item.attributes.caption).toBe(paragraphTextOf(part, native.caption!.id));
  return String(item.attributes.caption);
}
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const) {
  test(`grouped images seed one native caption owner before concurrent first edits ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({
      document: withoutNativeGalleryCaptions(await nativeGalleryDocument()),
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
      const native = readNativeGalleryImage(
        alice!.surface!.session.currentPackage(),
        alice!.surface!.session.part(),
        findNode(alice!.surface!.session.part(), gallery.images[0]!.controlId)!
      )!;
      expect(native.caption).not.toBeNull();
      expect(nativeCaption(alice!.surface!.session.currentPackage(), grouped.blockId)).toBe('');
      const captionId = native.caption!.id;
      pair.pause();
      for (const [at, editor] of pair.editors.entries())
        expect(
          updateGallery(editor!, {
            blockId: grouped.blockId,
            change: {
              type: 'caption',
              imageId: gallery.images[0]!.id,
              expectedCaption: '',
              caption: at ? 'Bob' : 'Alice',
            },
          }).success
        ).toBe(true);
      pair.reconnect();
      await pair.settle();
      expect(nativeCaption(alice!.surface!.session.currentPackage(), grouped.blockId)).toBe(
        nativeCaption(bob!.surface!.session.currentPackage(), grouped.blockId)
      );
      expect(findNode(alice!.surface!.session.part(), captionId)?.kind).toBe('paragraph');
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(nativeCaption(bob!.surface!.session.currentPackage(), grouped.blockId)).toBe('Bob');
      expect(nativeCaption(openGallery(new Uint8Array(await bob!.save())), grouped.blockId)).toBe(
        'Bob'
      );
    } finally {
      pair.destroy();
    }
  });
  test(`native gallery captions keep character ownership, native breaks, and history ${clientIds}`, async () => {
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
      const before = readPayloadRecords(alice!.surface!.session.currentPackage());
      pair.pause();
      expect(
        updateGallery(alice!, {
          blockId: grouped.blockId,
          change: {
            type: 'caption',
            imageId: gallery.images[0]!.id,
            expectedCaption: 'First image',
            caption: 'First\r\nAlice 🖼',
          },
        }).success
      ).toBe(true);
      expect(
        updateGallery(bob!, {
          blockId: grouped.blockId,
          change: {
            type: 'caption',
            imageId: gallery.images[1]!.id,
            expectedCaption: 'Second image',
            caption: 'Second\nBob',
          },
        }).success
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage();
        expect(nativeCaption(pkg, grouped.blockId)).toBe('First\nAlice 🖼');
        expect(nativeCaption(pkg, grouped.blockId, 1)).toBe('Second\nBob');
        expect(readPayloadRecords(pkg)).toEqual(before);
        const reopened = openGallery(new Uint8Array(await editor.save()));
        expect(nativeCaption(reopened, grouped.blockId)).toBe('First\nAlice 🖼');
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(nativeCaption(bob!.surface!.session.currentPackage(), grouped.blockId)).toBe(
        'First image'
      );
      expect(nativeCaption(bob!.surface!.session.currentPackage(), grouped.blockId, 1)).toBe(
        'Second\nBob'
      );
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      pair.pause();
      for (const [at, editor] of pair.editors.entries())
        expect(
          updateGallery(editor!, {
            blockId: grouped.blockId,
            change: {
              type: 'caption',
              imageId: gallery.images[0]!.id,
              expectedCaption: 'First\nAlice 🖼',
              caption: at === 0 ? 'First\nAlice 🖼 A' : 'First\nAlice 🖼 B',
            },
          }).success
        ).toBe(true);
      pair.reconnect(1);
      await pair.settle();
      expect(nativeCaption(alice!.surface!.session.currentPackage(), grouped.blockId)).toBe(
        nativeCaption(bob!.surface!.session.currentPackage(), grouped.blockId)
      );
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(nativeCaption(bob!.surface!.session.currentPackage(), grouped.blockId)).toBe(
        'First\nAlice 🖼 B'
      );
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
    } finally {
      pair.destroy();
    }
  });

  test(`native gallery deletion keeps peer edits and resources through undo ${clientIds}`, async () => {
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
      const originalMedia = media(alice!.surface!.session.currentPackage());
      const originalCaption = String(gallery.images[0]!.attributes.caption);
      pair.pause();
      expect(
        updateGallery(alice!, {
          blockId: grouped.blockId,
          change: { type: 'remove', imageId: gallery.images[0]!.id },
        }).success
      ).toBe(true);
      expect(
        updateGallery(bob!, {
          blockId: grouped.blockId,
          change: {
            type: 'caption',
            imageId: gallery.images[0]!.id,
            expectedCaption: originalCaption,
            caption: 'Caption while deleted',
          },
        }).success
      ).toBe(true);
      const drawing = nativeGalleryDrawings(bob!.surface!.session.currentPackage())[0]!.drawing;
      expect(
        updateInlineImage(bob!, {
          drawingNodeId: drawing.id,
          expectedRevision: bob!.surface!.session.packageRevision(),
          change: { type: 'alt', alt: 'Alt while deleted' },
        }).success
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      for (const editor of pair.editors)
        expect(
          readGallery(editor.surface!.session.currentPackage(), grouped.blockId)!.images
        ).toHaveLength(2);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage();
        expect(nativeCaption(pkg, grouped.blockId)).toBe('Caption while deleted');
        expect(readGallery(pkg, grouped.blockId)!.images[0]!.attributes.alt).toBe(
          'Alt while deleted'
        );
        expect(media(pkg)).toEqual(originalMedia);
      }
      for (let left = 3; left > 0; left--) {
        const current = readGallery(alice!.surface!.session.currentPackage(), grouped.blockId)!;
        expect(
          updateGallery(alice!, {
            blockId: grouped.blockId,
            change: { type: 'remove', imageId: current.images[0]!.id },
          }).success
        ).toBe(true);
        await pair.settle();
      }
      for (const editor of pair.editors) {
        expect(readGallery(editor.surface!.session.currentPackage(), grouped.blockId)).toBeNull();
        expect(media(editor.surface!.session.currentPackage())).toEqual([]);
        const reopened = openGallery(new Uint8Array(await editor.save()));
        expect(readGallery(reopened, grouped.blockId)).toBeNull();
        expect(media(reopened)).toEqual([]);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(
        readGallery(bob!.surface!.session.currentPackage(), grouped.blockId)!.images
      ).toHaveLength(1);
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
    } finally {
      pair.destroy();
    }
  });
}
