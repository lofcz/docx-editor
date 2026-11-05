import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, semanticDigest } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { galleryDocument, openGallery } from '../../test/support/gallery';
import { collectBlockClipboardData, pasteWithBlocks } from '../react-blocks/clipboard';
import { buildBlockIndex } from './block-index';
import { duplicateBlock } from './duplicate';
import { readGallery, updateGallery } from './gallery';
import { removeBlock } from './remove';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`duplicated galleries keep separate image owners through shared edits and deletion ${clients}`, async () => {
    const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds: clients });
    try {
      const [alice, bob] = pair.editors;
      const original = readGallery(alice!.surface!.session.currentPackage(), 'gallery')!;
      let serial = 0;
      const copy = duplicateBlock(alice!, {
        controlId: original.controlId,
        allocateBlockId: () => `copy-${++serial}`,
      });
      if (!copy.success) throw Error(copy.error);
      await pair.settle();
      const copyId = copy.payloadIds.gallery!;
      const copied = readGallery(bob!.surface!.session.currentPackage(), copyId)!;
      expect(copied.images.map((item) => item.id)).not.toEqual(
        original.images.map((item) => item.id)
      );
      pair.pause();
      expect(
        updateGallery(alice!, {
          blockId: copyId,
          change: {
            type: 'caption',
            imageId: copied.images[0]!.id,
            caption: 'Copied caption',
            expectedCaption: 'First image',
          },
        }).success
      ).toBe(true);
      expect(
        updateGallery(bob!, {
          blockId: 'gallery',
          change: {
            type: 'caption',
            imageId: original.images[0]!.id,
            caption: 'Original caption',
            expectedCaption: 'First image',
          },
        }).success
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage();
        expect(readGallery(pkg, copyId)!.images[0]!.attributes.caption).toBe('Copied caption');
        expect(readGallery(pkg, 'gallery')!.images[0]!.attributes.caption).toBe('Original caption');
        expect(buildBlockIndex(pkg).diagnostics).toEqual([]);
      }
      expect(
        removeBlock(alice!, {
          controlId: readGallery(alice!.surface!.session.currentPackage(), copyId)!.controlId,
          mode: 'delete',
        }).success
      ).toBe(true);
      await pair.settle();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(
        readGallery(bob!.surface!.session.currentPackage(), copyId)!.images[0]!.attributes.caption
      ).toBe('Copied caption');
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = openGallery(new Uint8Array(await editor.save()));
        expect(readGallery(pkg, copyId)).toBeNull();
        expect(readGallery(pkg, 'gallery')!.images[0]!.attributes.caption).toBe('Original caption');
      }
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
    } finally {
      pair.destroy();
    }
  });

for (const privateData of [true, false])
  test(`gallery clipboard repairs declared owner references and keeps captions literal: private=${privateData}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: galleryDocument(),
    });
    try {
      const surface = editor.surface!;
      const image = readGallery(surface.session.currentPackage(), 'gallery')!.images[0]!;
      expect(
        updateGallery(editor, {
          blockId: 'gallery',
          change: {
            type: 'caption',
            imageId: image.id,
            expectedCaption: 'First image',
            caption: 'gallery',
          },
        }).success
      ).toBe(true);
      const ids = surface.session.paragraphIds();
      const end = ids[2]!;
      surface.setSelection({
        anchor: { paragraphId: ids[0]!, offset: 0 },
        head: { paragraphId: end, offset: paragraphTextOf(surface.session.part(), end)!.length },
      });
      const flavours = surface.copyFlavours();
      const data = collectBlockClipboardData(editor)!;
      expect(data.blocks).toHaveLength(4);
      const tail = { paragraphId: ids[3]!, offset: 4 };
      surface.setSelection({ anchor: tail, head: tail });
      const before = semanticDigest(surface.session.currentPackage().parts.values());
      expect(
        pasteWithBlocks(
          editor,
          { ...flavours, data: privateData ? data : null },
          { createId: (payload) => `copy-${payload!.id}` }
        )
      ).toBe(true);
      const copy = readGallery(surface.session.currentPackage(), 'copy-gallery')!;
      expect(copy.images).toHaveLength(3);
      expect(copy.images.every((item) => item.id.startsWith('copy-'))).toBe(true);
      expect(copy.images[0]!.attributes.caption).toBe('gallery');
      expect(buildBlockIndex(surface.session.currentPackage()).diagnostics).toEqual([]);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(semanticDigest(surface.session.currentPackage().parts.values())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(
        readGallery(openGallery(new Uint8Array(await editor.save())), 'copy-gallery')!.images
      ).toHaveLength(3);
    } finally {
      editor.destroy();
    }
  });
