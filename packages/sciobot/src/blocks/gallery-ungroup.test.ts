import { expect, test } from 'bun:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  semanticDigest,
  bodyStoryRoot,
  directParagraphProperties,
  paragraphTextOf,
  storyParagraphs,
  type ImageDecodePort,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { galleryDocument, openGallery } from '../../test/support/gallery';
import { buildBlockIndex } from './block-index';
import { readGallery, updateGallery } from './gallery';
import { ungroupGallery } from './gallery-ungroup';

const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decodePort: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const resolve = async () => ({ bytes: png, mime: 'image/png' as const });
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes].filter(([name]) => /(^|\/)media\//.test(name));
const nativeImages = (pkg: OoxmlPackage) =>
  [...buildBlockIndex(pkg).byId.values()].filter((block) => block.payload.type === 'inline_image');

function captions(pkg: OoxmlPackage) {
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  return storyParagraphs(bodyStoryRoot(part)!).filter((paragraph) =>
    directParagraphProperties(part, paragraph.id).some(
      (property) =>
        property.localName === 'pStyle' &&
        (property.attributes?.val ?? property.attributes?.['w:val']) === 'Caption'
    )
  );
}

test('ungroup preserves visible multiline captions through peer editing, history, and reopen', async () => {
  const pair = await createTwoEditorFixture({
    document: galleryDocument(),
    clientIds: [73, 41],
    editor: { imageDecodePort: decodePort },
  });
  try {
    const [alice, bob] = pair.editors;
    const imageId = readGallery(alice!.surface!.session.currentPackage(), 'gallery')!.images[0]!.id;
    expect(
      updateGallery(alice!, {
        blockId: 'gallery',
        change: {
          type: 'caption',
          imageId,
          expectedCaption: 'First image',
          caption: 'First\r\nSecond',
        },
      }).success
    ).toBe(true);
    await pair.settle();
    expect(
      (await ungroupGallery(alice!, { blockId: 'gallery', decodePort, resolve })).success
    ).toBe(true);
    await pair.settle();
    const caption = captions(bob!.surface!.session.currentPackage())[0]!;
    expect(paragraphTextOf(bob!.surface!.session.part(), caption.id)).toBe('First\nSecond');
    const edited = bob!.surface!.applyAutomationOps(
      () => [{ op: 'insertText', paragraphId: caption.id, offset: 12, text: ' peer' }],
      { kind: 'body' },
      undefined,
      false,
      true
    );
    expect(edited.rejected).toBe(false);
    await pair.settle();
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(
      readGallery(bob!.surface!.session.currentPackage(), 'gallery')!.images[0]!.attributes.caption
    ).toBe('First\r\nSecond');
    expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors) {
      const pkg = openGallery(new Uint8Array(await editor.save()));
      const part = pkg.parts.get(pkg.mainDocumentPart)!;
      expect(captions(pkg).map((paragraph) => paragraphTextOf(part, paragraph.id))).toEqual([
        'First\nSecond peer',
        'Second image',
        'Third image',
      ]);
    }
  } finally {
    pair.destroy();
  }
});

test('ungroup after appending retains every image beside a preceding paragraph', async () => {
  const files = unzipSync(galleryDocument());
  files['word/document.xml'] = strToU8(
    strFromU8(files['word/document.xml']!).replace(
      '<w:body>',
      '<w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>'
    )
  );
  const pair = await createTwoEditorFixture({
    document: zipSync(files),
    editor: { imageDecodePort: decodePort },
  });
  try {
    const alice = pair.editors[0]!;
    expect(
      updateGallery(alice, {
        blockId: 'gallery',
        change: {
          type: 'append',
          images: [
            {
              id: 'appended',
              attributes: { url: 'https://assets.example.test/added.png', width: 1, height: 1 },
            },
          ],
        },
      }).success
    ).toBe(true);
    await pair.settle();
    expect((await ungroupGallery(alice, { blockId: 'gallery', decodePort, resolve })).success).toBe(
      true
    );
    await pair.settle();
    for (const editor of pair.editors) {
      const pkg = editor.surface!.session.currentPackage();
      expect(buildBlockIndex(pkg).diagnostics).toEqual([]);
      expect(nativeImages(pkg)).toHaveLength(4);
    }
  } finally {
    pair.destroy();
  }
});

test('ungroup embeds distinct native images in one transaction and preserves history and DOCX', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: galleryDocument(),
    imageDecodePort: decodePort,
  });
  try {
    const session = editor.surface!.session;
    const original = semanticDigest(session.currentPackage().parts.values());
    const references: string[] = [];
    let publications = 0;
    const stop = session.subscribe(() => publications++);
    const result = await ungroupGallery(editor, {
      blockId: 'gallery',
      decodePort,
      resolve: async (reference) => {
        references.push(reference);
        return resolve();
      },
    });
    stop();
    expect(result.success).toBe(true);
    expect(publications).toBe(1);
    expect(references).toEqual([
      'https://assets.example.test/same.png',
      'https://assets.example.test/third.png',
    ]);
    const pkg = session.currentPackage();
    expect(readGallery(pkg, 'gallery')).toBeNull();
    const images = nativeImages(pkg);
    expect(images).toHaveLength(3);
    expect(new Set(images.map((image) => image.id)).size).toBe(3);
    expect(images.map((image) => image.payload.attributes.alt)).toEqual([
      'First image',
      'Second image',
      'Third image',
    ]);
    expect(images[0]!.payload.attributes.widthPoints).toBe(48);
    expect(images[0]!.payload.attributes.heightPoints).toBe(24);
    expect(captions(pkg).map((paragraph) => paragraphTextOf(session.part(), paragraph.id))).toEqual(
      ['First image', 'Second image', 'Third image']
    );
    expect(media(pkg)).toHaveLength(3);
    for (const [, bytes] of media(pkg)) expect(bytes).toEqual(png);
    expect(buildBlockIndex(pkg).diagnostics).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(semanticDigest(session.currentPackage().parts.values())).toEqual(original);
    expect(media(session.currentPackage())).toEqual([]);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const reopened = openGallery(new Uint8Array(await editor.save()));
    expect(nativeImages(reopened).map((image) => image.payload)).toEqual(
      images.map((image) => image.payload)
    );
    expect(media(reopened)).toEqual(media(pkg));
  } finally {
    editor.destroy();
  }
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`ungroup and peer captions converge and survive both histories and reconnect ${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      document: galleryDocument(),
      clientIds: clients,
      editor: { imageDecodePort: decodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const imageId = readGallery(bob!.surface!.session.currentPackage(), 'gallery')!.images[0]!.id;
      pair.pause();
      expect(
        (await ungroupGallery(alice!, { blockId: 'gallery', resolve, decodePort })).success
      ).toBe(true);
      expect(
        updateGallery(bob!, {
          blockId: 'gallery',
          change: {
            type: 'caption',
            imageId,
            expectedCaption: 'First image',
            caption: 'Peer caption',
          },
        }).success
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      for (const editor of pair.editors) {
        expect(readGallery(editor.surface!.session.currentPackage(), 'gallery')).toBeNull();
        expect(nativeImages(editor.surface!.session.currentPackage())).toHaveLength(3);
        expect(buildBlockIndex(editor.surface!.session.currentPackage()).diagnostics).toEqual([]);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(
          readGallery(editor.surface!.session.currentPackage(), 'gallery')!.images[0]!.attributes
            .caption
        ).toBe('Peer caption');
        expect(nativeImages(editor.surface!.session.currentPackage())).toHaveLength(0);
      }
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(
        readGallery(alice!.surface!.session.currentPackage(), 'gallery')!.images[0]!.attributes
          .caption
      ).toBe('First image');
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      pair.pause();
      pair.reconnect(1);
      await pair.settle();
      for (const editor of pair.editors) {
        const reopened = openGallery(new Uint8Array(await editor.save()));
        expect(readGallery(reopened, 'gallery')).toBeNull();
        expect(nativeImages(reopened)).toHaveLength(3);
        expect(buildBlockIndex(reopened).diagnostics).toEqual([]);
      }
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  });

test('ungroup refuses failed assets and stale asynchronous preparation without partial writes', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: galleryDocument(),
  });
  try {
    const session = editor.surface!.session;
    const before = session.currentPackage();
    const failed = await ungroupGallery(editor, {
      blockId: 'gallery',
      decodePort,
      resolve: async () => {
        throw Error('Asset unavailable');
      },
    });
    expect(failed.success).toBe(false);
    expect(session.currentPackage()).toBe(before);
    let changed = false;
    const stale = await ungroupGallery(editor, {
      blockId: 'gallery',
      decodePort,
      resolve: async () => {
        if (!changed) {
          changed = true;
          const image = readGallery(session.currentPackage(), 'gallery')!.images[0]!;
          expect(
            updateGallery(editor, {
              blockId: 'gallery',
              change: {
                type: 'caption',
                imageId: image.id,
                expectedCaption: 'First image',
                caption: 'Changed',
              },
            }).success
          ).toBe(true);
        }
        return resolve();
      },
    });
    expect(stale).toMatchObject({ success: false, code: 'stale-revision' });
    expect(readGallery(session.currentPackage(), 'gallery')!.images[0]!.attributes.caption).toBe(
      'Changed'
    );
    expect(nativeImages(session.currentPackage())).toHaveLength(0);
    expect(media(session.currentPackage())).toHaveLength(0);
  } finally {
    editor.destroy();
  }
});
