import { expect, test } from 'bun:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { galleryDocument, openGallery } from '../../test/support/gallery';
import {
  findNode,
  contentControlContentNodeOf,
  paragraphTextOf,
  storyParagraphs,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { readGallery, updateGallery } from './gallery';

for (const lock of ['contentLocked', 'sdtContentLocked', 'sdtLocked'] as const)
  test(`gallery edits honor external ancestor ${lock} without changing native lock declarations`, async () => {
    const files = unzipSync(galleryDocument());
    files['word/document.xml'] = strToU8(
      strFromU8(files['word/document.xml']!)
        .replace(
          '<w:body>',
          `<w:body><w:sdt><w:sdtPr><w:lock w:val="${lock}"/></w:sdtPr><w:sdtContent>`
        )
        .replace('</w:body>', '</w:sdtContent></w:sdt></w:body>')
    );
    const pair = await createTwoEditorFixture({ document: zipSync(files) });
    try {
      const editor = pair.editors[0]!;
      const imageId = readGallery(editor.surface!.session.currentPackage(), 'gallery')!.images[0]!
        .id;
      const before = editor.surface!.session.currentPackage();
      const permitted = lock === 'sdtLocked';
      expect(
        updateGallery(editor, {
          blockId: 'gallery',
          change: { type: 'caption', imageId, expectedCaption: 'First image', caption: 'Caption' },
        })
      ).toMatchObject({ success: permitted });
      expect(
        updateGallery(editor, {
          blockId: 'gallery',
          change: { type: 'presentation', layout: 'grid' },
        }).success
      ).toBe(permitted);
      expect(
        updateGallery(editor, {
          blockId: 'gallery',
          change: {
            type: 'append',
            images: [
              { id: 'added-image', attributes: { url: 'https://assets.example.test/added.png' } },
            ],
          },
        }).success
      ).toBe(permitted);
      await pair.settle();
      if (!permitted) expect(editor.surface!.session.currentPackage()).toBe(before);
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  });

test('multiline gallery captions retain exact payloads and stable readable fallback through DOCX and shared history', async () => {
  const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds: [73, 41] });
  try {
    const editor = pair.editors[0]!;
    const imageId = readGallery(editor.surface!.session.currentPackage(), 'gallery')!.images[0]!.id;
    const caption = 'Line one\r\nLine two\rLine three\nLine four';
    expect(
      updateGallery(editor, {
        blockId: 'gallery',
        change: { type: 'caption', imageId, expectedCaption: 'First image', caption },
      })
    ).toMatchObject({ success: true });
    await pair.settle();
    for (const member of pair.editors) {
      for (const pkg of [
        member.surface!.session.currentPackage(),
        openGallery(new Uint8Array(await member.save())),
      ]) {
        const gallery = readGallery(pkg, 'gallery')!;
        expect(gallery.images[0]!.attributes.caption).toBe(caption);
        const part = pkg.parts.get(gallery.partName)!;
        const paragraph = storyParagraphs(
          contentControlContentNodeOf(findNode(part, gallery.images[0]!.controlId)!)!
        )[0]!;
        expect(paragraphTextOf(part, paragraph.id)).toBe('Line one Line two Line three Line four');
      }
    }
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(
      readGallery(pair.editors[1]!.surface!.session.currentPackage(), 'gallery')!.images[0]!
        .attributes.caption
    ).toBe('First image');
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    expect(
      semanticDigest(pair.editors[0]!.surface!.session.currentPackage().parts.values())
    ).toEqual(semanticDigest(pair.editors[1]!.surface!.session.currentPackage().parts.values()));
  } finally {
    pair.destroy();
  }
});
