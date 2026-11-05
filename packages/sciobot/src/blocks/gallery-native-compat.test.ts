import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  findNode,
  paragraphTextOf,
  semanticDigest,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { decodeBlockPayload as oldDecode } from '../../test/compat/payload-reader-v2';
import {
  nativeGalleryDocument,
  nativeGalleryDrawings,
  galleryDecodePort,
} from '../../test/support/native-gallery';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { openGallery } from '../../test/support/gallery';
import { groupAdjacentImages } from './gallery-group';
import { readGallery } from './gallery';
import { readPayloadRecords, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { updateAtomicBlock } from './update-atomic';
import { readNativeGalleryImage } from './gallery-native-record';

test('version-two readers keep native galleries opaque with unchanged native images and visible captions', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: await nativeGalleryDocument('First\r\nSecond'),
    imageDecodePort: galleryDecodePort,
  });
  try {
    const surface = editor.surface!;
    const original = surface.session.currentPackage();
    const grouped = groupAdjacentImages(editor, {
      drawingNodeId: nativeGalleryDrawings(original)[0]!.drawing.id,
    });
    if (!grouped.success) throw Error(grouped.error);
    const pkg = surface.session.currentPackage();
    const records = readPayloadRecords(pkg);
    if (records.status !== 'ready') throw Error('Missing native gallery records');
    const nativeRecords = [...records.entries.values()].filter(
      (record) => record.decoded.status === 'supported' && record.decoded.value.version === 3
    );
    expect(nativeRecords).toHaveLength(4);
    for (const record of nativeRecords) {
      const old = oldDecode(record.raw);
      expect(old.status).toBe('unsupported-version');
      expect('raw' in old && old.raw).toBe(record.raw);
    }
    const reopened = openGallery(writeOoxmlPackage(pkg));
    expect(semanticDigest(reopened.parts.values())).toEqual(semanticDigest(pkg.parts.values()));
    const store = readPayloadRecords(reopened);
    if (store.status !== 'ready') throw Error('Missing reopened metadata');
    expect([...store.entries.values()].map((record) => record.raw)).toEqual(
      [...records.entries.values()].map((record) => record.raw)
    );
    const part = reopened.parts.get(reopened.mainDocumentPart)!;
    expect(
      storyParagraphs(bodyStoryRoot(part)!).map((paragraph) => paragraphTextOf(part, paragraph.id))
    ).toEqual([
      '\uFFFC',
      'First\nSecond',
      '\uFFFC',
      'Second image',
      '\uFFFC',
      'Third image',
      'Tail',
    ]);
    expect([...reopened.partBytes].filter(([name]) => name.includes('/media/'))).toEqual(
      [...original.partBytes].filter(([name]) => name.includes('/media/'))
    );
    const gallery = readGallery(pkg, grouped.blockId)!;
    expect(
      updateAtomicBlock(editor, {
        blockId: grouped.blockId,
        attributes: gallery.payload.attributes,
      })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(
      withPayloadRecord(
        pkg,
        { ...gallery.payload, attributes: { ...gallery.payload.attributes, images: [] } },
        'Gallery'
      )
    ).toMatchObject({ ok: false, reason: 'gallery-images-require-owned-record-command' });
    const projected = readPayloadStore(pkg);
    if (projected.status !== 'ready') throw Error('Missing projected metadata');
    const item = projected.entries.get(gallery.images[0]!.id)!;
    if (item.decoded.status !== 'supported') throw Error('Missing image owner');
    expect(
      withPayloadRecord(
        pkg,
        {
          ...item.decoded.value,
          attributes: { ...item.decoded.value.attributes, alt: 'Wrong authority' },
        },
        item.label
      )
    ).toMatchObject({ ok: false, reason: 'gallery-image-fields-require-native-command' });
    expect(
      readNativeGalleryImage(
        pkg,
        surface.session.part(),
        findNode(surface.session.part(), gallery.images[0]!.controlId)!
      )!.attributes
    ).toMatchObject({ alt: 'First\nSecond', caption: 'First\nSecond' });
    expect(surface.session.currentPackage()).toBe(pkg);
  } finally {
    editor.destroy();
  }
});

for (const lock of ['contentLocked', 'sdtContentLocked'])
  test(`native grouping respects the enclosing ${lock} owner`, async () => {
    const files = unzipSync(await nativeGalleryDocument());
    files['word/document.xml'] = strToU8(
      strFromU8(files['word/document.xml']!)
        .replace(
          '<w:body>',
          `<w:body><w:sdt><w:sdtPr><w:lock w:val="${lock}"/></w:sdtPr><w:sdtContent>`
        )
        .replace('</w:body>', '</w:sdtContent></w:sdt></w:body>')
    );
    const pair = await createTwoEditorFixture({
      document: zipSync(files),
      ownedAttributes: true,
      editor: { imageDecodePort: galleryDecodePort },
    });
    try {
      const editor = pair.editors[0]!;
      const pkg = editor.surface!.session.currentPackage();
      const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      expect(
        groupAdjacentImages(editor, { drawingNodeId: nativeGalleryDrawings(pkg)[0]!.drawing.id })
      ).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(editor.surface!.session.currentPackage()).toBe(pkg);
      expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
    } finally {
      pair.destroy();
    }
  });

test('grouping refuses an unsupported shared history scope without changing canonical or shared bytes', async () => {
  const pair = await createTwoEditorFixture({
    document: await nativeGalleryDocument(),
    editor: { imageDecodePort: galleryDecodePort },
  });
  try {
    const editor = pair.editors[0]!;
    const pkg = editor.surface!.session.currentPackage();
    const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    expect(
      groupAdjacentImages(editor, { drawingNodeId: nativeGalleryDrawings(pkg)[0]!.drawing.id })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(editor.surface!.session.currentPackage()).toBe(pkg);
    expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
  } finally {
    pair.destroy();
  }
});
