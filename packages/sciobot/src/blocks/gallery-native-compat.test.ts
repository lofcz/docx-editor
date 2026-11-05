import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
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
import { readGallery, updateGallery } from './gallery';
import { ungroupGallery } from './gallery-ungroup';
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

test('recorded Runtime G reader and writer preserve version-three gallery bytes and native fallback', () => {
  const fixture = (name: string) =>
    readFileSync(new URL(`../../test/compat/fixtures/${name}`, import.meta.url));
  const input = fixture('native-gallery-v3-input.docx');
  const output = fixture('native-gallery-v3-old-output.docx');
  const provenance = JSON.parse(fixture('native-gallery-v3-old-writer.json').toString());
  expect(provenance.revision).toBe('c3bf6048a');
  expect(createHash('sha256').update(input).digest('hex')).toBe(provenance.inputSha256);
  expect(createHash('sha256').update(output).digest('hex')).toBe(provenance.outputSha256);
  const before = openGallery(input),
    after = openGallery(output);
  expect(semanticDigest(after.parts.values())).toEqual(semanticDigest(before.parts.values()));
  expect([...after.partBytes].filter(([name]) => name.includes('/media/'))).toEqual(
    [...before.partBytes].filter(([name]) => name.includes('/media/'))
  );
  const beforeRecords = readPayloadRecords(before),
    afterRecords = readPayloadRecords(after);
  if (beforeRecords.status !== 'ready' || afterRecords.status !== 'ready')
    throw Error('Missing old-writer gallery records');
  expect([...afterRecords.entries.values()].map((record) => record.raw)).toEqual(
    [...beforeRecords.entries.values()].map((record) => record.raw)
  );
  const nativeRecords = [...afterRecords.entries.values()].filter(
    (record) => record.decoded.status === 'supported' && record.decoded.value.version === 3
  );
  expect(nativeRecords).toHaveLength(4);
  for (const record of nativeRecords)
    expect(oldDecode(record.raw).status).toBe('unsupported-version');
});

test('a body image request refuses while the native header owns editing', async () => {
  const files = unzipSync(await nativeGalleryDocument());
  const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  files['[Content_Types].xml'] = strToU8(
    strFromU8(files['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
    )
  );
  files['word/_rels/document.xml.rels'] = strToU8(
    strFromU8(files['word/_rels/document.xml.rels']!).replace(
      '</Relationships>',
      '<Relationship Id="nativeGalleryHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'
    )
  );
  files['word/header1.xml'] = strToU8(
    `<w:hdr xmlns:w="${word}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`
  );
  files['word/document.xml'] = strToU8(
    strFromU8(files['word/document.xml']!).replace(
      '</w:body>',
      '<w:sectPr><w:headerReference w:type="default" r:id="nativeGalleryHeader"/></w:sectPr></w:body>'
    )
  );
  const pair = await createTwoEditorFixture({
    document: zipSync(files),
    ownedAttributes: true,
    editor: { imageDecodePort: galleryDecodePort },
  });
  try {
    const editor = pair.editors[0]!;
    const drawing = nativeGalleryDrawings(editor.surface!.session.currentPackage())[0]!.drawing;
    expect(editor.surface!.enterHeaderFooter({ rId: 'nativeGalleryHeader', pageIndex: 0 })).toBe(
      true
    );
    expect(editor.surface!.activeScope().kind).not.toBe('body');
    const pkg = editor.surface!.session.currentPackage();
    const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    expect(groupAdjacentImages(editor, { drawingNodeId: drawing.id })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    expect(editor.surface!.session.currentPackage()).toBe(pkg);
    expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
    editor.surface!.exitHeaderFooter();
    const grouped = groupAdjacentImages(editor, { drawingNodeId: drawing.id });
    if (!grouped.success) throw Error(grouped.error);
    await pair.settle();
    expect(editor.surface!.enterHeaderFooter({ rId: 'nativeGalleryHeader', pageIndex: 0 })).toBe(
      true
    );
    const groupedPackage = editor.surface!.session.currentPackage();
    const groupedShared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    expect(
      updateGallery(editor, {
        blockId: grouped.blockId,
        change: { type: 'presentation', layout: 'grid' },
      })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(
      await ungroupGallery(editor, {
        blockId: grouped.blockId,
        decodePort: galleryDecodePort,
        resolve: async () => {
          throw Error('No resource access is permitted');
        },
      })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(editor.surface!.session.currentPackage()).toBe(groupedPackage);
    expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(groupedShared);
  } finally {
    pair.destroy();
  }
});

for (const at of [0, 1, 2])
  for (const lock of ['contentLocked', 'sdtLocked'])
    test(`grouping honors managed picture ${at} with ${lock}`, async () => {
      const files = unzipSync(await nativeGalleryDocument());
      let count = 0;
      files['word/document.xml'] = strToU8(
        strFromU8(files['word/document.xml']!).replace(
          /<w:lock w:val="sdtContentLocked"\s*\/>/g,
          (source) => (count++ === at ? `<w:lock w:val="${lock}"/>` : source)
        )
      );
      expect(count).toBe(3);
      const pair = await createTwoEditorFixture({
        document: zipSync(files),
        ownedAttributes: true,
        editor: { imageDecodePort: galleryDecodePort },
      });
      try {
        const editor = pair.editors[0]!,
          pkg = editor.surface!.session.currentPackage();
        const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
        expect(
          groupAdjacentImages(editor, { drawingNodeId: nativeGalleryDrawings(pkg)[at]!.drawing.id })
        ).toMatchObject({ success: false, code: 'unsupported-operation' });
        expect(editor.surface!.session.currentPackage()).toBe(pkg);
        expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
      } finally {
        pair.destroy();
      }
    });

for (const at of [0, 1, 2])
  test(`grouping keeps protected caption ${at} with its picture`, async () => {
    const files = unzipSync(await nativeGalleryDocument());
    const caption = ['First image', 'Second image', 'Third image'][at]!;
    const source = strFromU8(files['word/document.xml']!);
    const expression = new RegExp(`<w:r>(<w:t[^>]*>${caption}</w:t>)</w:r>`);
    expect(expression.test(source)).toBe(true);
    files['word/document.xml'] = strToU8(
      source.replace(
        expression,
        '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:r>$1</w:r></w:sdtContent></w:sdt>'
      )
    );
    const pair = await createTwoEditorFixture({
      document: zipSync(files),
      ownedAttributes: true,
      editor: { imageDecodePort: galleryDecodePort },
    });
    try {
      const editor = pair.editors[0]!,
        pkg = editor.surface!.session.currentPackage();
      const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      expect(
        groupAdjacentImages(editor, { drawingNodeId: nativeGalleryDrawings(pkg)[at]!.drawing.id })
      ).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(editor.surface!.session.currentPackage()).toBe(pkg);
      expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
    } finally {
      pair.destroy();
    }
  });
