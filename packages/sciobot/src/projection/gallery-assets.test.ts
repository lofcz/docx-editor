import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { openGallery } from '../../test/support/gallery';
import {
  galleryDecodePort,
  galleryPng,
  nativeGalleryDocument,
  nativeGalleryDrawings,
} from '../../test/support/native-gallery';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from './revision-cache';
import { galleryAssetPayload } from './native-gallery';
import { imageAssetBundleResolver } from '../native/image-asset-bundle';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { groupAdjacentImages } from '../blocks/gallery-group';
import { readGallery, updateGallery } from '../blocks/gallery';
import { readGalleryImageAsset } from '../blocks/gallery-native-asset';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadRecords } from '../blocks/payload-opc-codec';
import { planTargetDocument } from '../diff/plan-target';
import { applyDirectPlan } from '../ai/apply-direct';
import { removeBlock } from '../blocks/remove';
import * as Y from 'yjs';

const context = { documentId: 'gallery-assets', epoch: 'test' };
async function view(editor: DocxEditorInstance, includeImageAssets = false) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const result = await readWorkingView(editor.surface!.session, revision, { includeImageAssets });
  if (!result.success) throw Error(result.error);
  return result.view;
}
async function exported() {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: await nativeGalleryDocument('First\ncaption'),
    imageDecodePort: galleryDecodePort,
  });
  try {
    const grouped = groupAdjacentImages(editor, {
      drawingNodeId: nativeGalleryDrawings(editor.surface!.session.currentPackage())[0]!.drawing.id,
    });
    if (!grouped.success) throw Error(grouped.error);
    const gallery = readGallery(editor.surface!.session.currentPackage(), grouped.blockId)!;
    expect(
      updateGallery(editor, {
        blockId: grouped.blockId,
        change: {
          type: 'move',
          imageId: gallery.images[2]!.id,
          beforeImageId: gallery.images[1]!.id,
          expectedOrder: gallery.images.map((image) => image.id),
        },
      }).success
    ).toBe(true);
    expect(
      updateGallery(editor, {
        blockId: grouped.blockId,
        change: { type: 'presentation', layout: 'grid', align: 'right' },
      }).success
    ).toBe(true);
    const before = editor.surface!.session.currentPackage();
    const result = await view(editor, true);
    const assets = result.media!;
    const payload = galleryAssetPayload(before, grouped.blockId);
    expect(editor.surface!.session.currentPackage()).toBe(before);
    expect(assets).toHaveLength(1);
    expect(assets[0]!.occurrences).toHaveLength(3);
    expect(assets[0]!.bytes).toEqual(galleryPng);
    expect(result.projection.text).toContain('resourceVersion="1"');
    expect(result.projection.text).not.toMatch(/blob:|data:|https?:/);
    return { markdown: result.projection.text, assets, attributes: payload.attributes };
  } finally {
    editor.destroy();
  }
}
async function planned(editor: DocxEditorInstance, target: string) {
  const source = await view(editor);
  const result = planTargetDocument(
    {
      revision: source.revision,
      pkg: editor.surface!.session.currentPackage(),
      projection: source.projection,
    },
    target
  );
  if (!result.success) throw Error(result.error);
  return result.plan;
}
function galleries(pkg: OoxmlPackage) {
  return [...buildBlockIndex(pkg).byId.values()].filter(
    (block) => block.payload.type === 'image_gallery'
  );
}
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`gallery attachment bundle reimports native owners, peer edits, history and DOCX ${clientIds}`, async () => {
    const source = await exported();
    const pair = await createTwoEditorFixture({
      document: docxFixture('<w:p><w:r><w:t>Tail</w:t></w:r></w:p>'),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: galleryDecodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const plan = await planned(alice!, source.markdown);
      const prepared = await prepareInlineImageResources(
        plan,
        imageAssetBundleResolver(source.assets),
        galleryDecodePort
      );
      if (!prepared.success) throw Error(prepared.error);
      const before = alice!.surface!.session.currentPackage();
      expect(applyDirectPlan(alice!, plan, context).success).toBe(false);
      expect(alice!.surface!.session.currentPackage()).toBe(before);
      pair.pause();
      const applied = applyDirectPlan(alice!, plan, { ...context, images: prepared.images });
      if (!applied.success) throw Error(applied.error);
      const tail = bob!.surface!.session.paragraphIds()[0]!;
      expect(
        bob!.surface!.applyAutomationOps(
          () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' peer' }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).rejected
      ).toBe(false);
      pair.reconnect();
      await pair.settle();
      const galleryId = galleries(alice!.surface!.session.currentPackage())[0]!.id;
      for (const editor of pair.editors) {
        const pkg = editor.surface!.session.currentPackage();
        expect(galleries(pkg)).toHaveLength(1);
        expect(galleryAssetPayload(pkg, galleryId).attributes).toEqual(source.attributes);
        const gallery = readGallery(pkg, galleryId)!;
        expect(new Set(gallery.images.map((image) => image.id)).size).toBe(3);
        for (const image of gallery.images)
          expect(readGalleryImageAsset(pkg, galleryId, image.id)!.bytes).toEqual(galleryPng);
        const records = readPayloadRecords(pkg);
        if (records.status !== 'ready') throw Error('No gallery records');
        for (const image of gallery.images) {
          const record = records.entries.get(image.id)!.decoded;
          if (record.status !== 'supported') throw Error('Unsupported gallery item');
          expect(record.value.attributes).toEqual({ resource: { version: 1, kind: 'embedded' } });
        }
        expect((await view(editor)).projection.text).toBe(
          source.markdown.replace(/Tail$/, 'Tail peer')
        );
        expect((await planned(editor, (await view(editor)).projection.text)).operations).toEqual(
          []
        );
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(galleries(bob!.surface!.session.currentPackage())).toHaveLength(0);
      expect(paragraphTextOf(bob!.surface!.session.part(), tail)).toBe('Tail peer');
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const first = readGallery(bob!.surface!.session.currentPackage(), galleryId)!.images[0]!;
      expect(
        updateGallery(bob!, {
          blockId: galleryId,
          change: {
            type: 'caption',
            imageId: first.id,
            expectedCaption: 'First\ncaption',
            caption: 'Peer\ncaption',
          },
        }).success
      ).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = openGallery(new Uint8Array(await editor.save()));
        expect(
          readGallery(pkg, galleryId)!.images.map((image) => image.attributes.caption)
        ).toEqual(['Peer\ncaption', 'Third image', 'Second image']);
        for (const image of readGallery(pkg, galleryId)!.images)
          expect(readGalleryImageAsset(pkg, galleryId, image.id)!.bytes).toEqual(galleryPng);
      }
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
      pair.pause();
      expect(
        removeBlock(alice!, {
          controlId: readGallery(alice!.surface!.session.currentPackage(), galleryId)!.controlId,
          mode: 'delete',
        }).success
      ).toBe(true);
      expect(
        updateGallery(bob!, {
          blockId: galleryId,
          change: {
            type: 'caption',
            imageId: first.id,
            expectedCaption: 'Peer\ncaption',
            caption: 'Caption during deletion',
          },
        }).success
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = openGallery(new Uint8Array(await editor.save()));
        expect(readGallery(pkg, galleryId)).toBeNull();
        expect([...pkg.partBytes.keys()].filter((name) => name.includes('/media/'))).toEqual([]);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const pkg = openGallery(new Uint8Array(await editor.save()));
        expect(readGallery(pkg, galleryId)!.images[0]!.attributes.caption).toBe(
          'Caption during deletion'
        );
        for (const image of readGallery(pkg, galleryId)!.images)
          expect(readGalleryImageAsset(pkg, galleryId, image.id)!.bytes).toEqual(galleryPng);
      }
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
    } finally {
      pair.destroy();
    }
  });

test('attachment admission refuses missing, duplicate, mismatched and hostile inputs before mutation', async () => {
  const source = await exported();
  const pair = await createTwoEditorFixture({
    document: docxFixture('<w:p><w:r><w:t>Tail</w:t></w:r></w:p>'),
    ownedAttributes: true,
    editor: { imageDecodePort: galleryDecodePort },
  });
  try {
    const editor = pair.editors[0]!,
      plan = await planned(editor, source.markdown);
    const pkg = editor.surface!.session.currentPackage(),
      fingerprint = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    const asset = source.assets[0]!;
    for (const bundle of [
      [asset, asset],
      [{ ...asset, path: '../secret.png' }],
      [{ ...asset, path: 'blob:secret' }],
      [{ ...asset, id: '0'.repeat(64) }],
      [{ ...asset, bytes: new Uint8Array([1, 2]) }],
    ])
      expect(() => imageAssetBundleResolver(bundle)).toThrow();
    expect(
      (await prepareInlineImageResources(plan, imageAssetBundleResolver([]), galleryDecodePort))
        .success
    ).toBe(false);
    expect(
      (
        await prepareInlineImageResources(
          plan,
          async () => ({ bytes: galleryPng, mime: 'image/png' }),
          galleryDecodePort,
          AbortSignal.abort()
        )
      ).success
    ).toBe(false);
    for (const path of [
      '../secret.png',
      'https://example.test/image.png',
      'blob:secret',
      'media/' + '0'.repeat(64) + '.png',
    ]) {
      const changed = source.markdown.replaceAll(asset.path, path);
      const read = await view(editor);
      const result = planTargetDocument(
        { revision: read.revision, pkg, projection: read.projection },
        changed
      );
      if (path.startsWith('media/')) {
        if (!result.success) throw Error(result.error);
        expect(
          (
            await prepareInlineImageResources(
              result.plan,
              async () => ({ bytes: galleryPng, mime: 'image/png' }),
              galleryDecodePort
            )
          ).success
        ).toBe(false);
      } else expect(result.success).toBe(false);
    }
    expect(editor.surface!.session.currentPackage()).toBe(pkg);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
    expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
  } finally {
    pair.destroy();
  }
});
