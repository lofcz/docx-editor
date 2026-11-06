import { expect, test } from 'bun:test';
import { runInNewContext } from 'node:vm';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  galleryDecodePort,
  galleryPng,
  nativeGalleryDocument,
  nativeGalleryDrawings,
} from '../../test/support/native-gallery';
import { groupAdjacentImages } from '../blocks/gallery-group';
import { readGallery } from '../blocks/gallery';
import { readEditorRevision } from './revision-cache';
import { readWorkingView } from '../review/working-view';
import { readCommittedView } from '../review/committed-view';
import { planTargetDocument } from '../diff/plan-target';
import {
  prepareInlineImageResources,
  resolveInlineImageResources,
} from '../native/inline-image-resources';
import { imageAssetBundleResolver } from '../native/image-asset-bundle';
import { applyDirectPlan } from '../ai/apply-direct';
import { createBlockPresenceRecord } from '../review/block-presence-plan';
import { withBlockPresenceRecord } from '../review/block-presence-store';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { sha256Bytes } from '@docx-editor.dev/core/store';
import * as Y from 'yjs';

const context = {
  documentId: 'gallery-boundaries',
  epoch: 'test',
  author: 'Gallery author',
  date: '2026-11-06T00:00:00.000Z',
};
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`existing native gallery syntax edits refuse and exact review exports isolate bytes ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await nativeGalleryDocument(),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: galleryDecodePort, modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const editor = pair.editors[0]!;
      const grouped = groupAdjacentImages(editor, {
        drawingNodeId: nativeGalleryDrawings(editor.surface!.session.currentPackage())[0]!.drawing
          .id,
      });
      if (!grouped.success) throw Error(grouped.error);
      await pair.settle();
      const session = editor.surface!.session,
        pkg = session.currentPackage();
      const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const source = await readWorkingView(session, revision, { includeImageAssets: true });
      if (!source.success) throw Error(source.error);
      const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      for (const target of [
        source.view.projection.text.replace('First image', 'Edited caption'),
        source.view.projection.text.replace('"width":64', '"width":128'),
        source.view.projection.text.replace('layout="row-fit"', 'layout="grid"'),
      ]) {
        expect(target).not.toBe(source.view.projection.text);
        const plan = planTargetDocument(
          { revision, pkg, projection: source.view.projection },
          target
        );
        if (plan.success) {
          const resources = await prepareInlineImageResources(
            plan.plan,
            imageAssetBundleResolver(source.view.media!),
            galleryDecodePort
          );
          if (!resources.success) throw Error(resources.error);
          const result = applyDirectPlan(editor, plan.plan, {
            ...context,
            images: resources.images,
          });
          expect(result).toMatchObject({ success: false, code: 'unsupported-operation' });
        } else expect(plan.code).toBe('unsupported-operation');
        expect(session.currentPackage()).toBe(pkg);
        expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
      }
      const gallery = readGallery(pkg, grouped.blockId)!;
      const removal = createBlockPresenceRecord(
        pkg,
        pkg,
        [{ blockId: gallery.id, before: { present: true }, after: { present: false } }],
        { id: 'remove-gallery', author: context.author, date: context.date }
      );
      if (!removal.success) throw Error(removal.error);
      expect(
        editor.surface!.applyAutomationOps(
          () => [],
          { kind: 'body' },
          [
            (working) => {
              const saved = withBlockPresenceRecord(working, removal.record);
              if (!saved.success) throw Error(saved.error);
              return saved.pkg;
            },
          ],
          true,
          true
        ).rejected
      ).toBe(false);
      await pair.settle();
      for (const peer of pair.editors) {
        const token = readEditorRevision(peer, context.documentId, context.epoch)!;
        const snapshot = peer.surface!.session.currentPackage();
        const committed = await readCommittedView(peer.surface!.session, token, {
          includeImageAssets: true,
        });
        const working = await readWorkingView(peer.surface!.session, token, {
          includeImageAssets: true,
        });
        if (!committed.success || !working.success)
          throw Error(JSON.stringify({ committed, working }));
        expect(committed.view.media).toHaveLength(1);
        expect(committed.view.media![0]!.bytes).toEqual(galleryPng);
        expect(committed.view.projection.text).toContain('resourceVersion="1"');
        expect(working.view.media).toEqual([]);
        expect(working.view.projection.text).not.toContain('image_gallery');
        committed.view.media![0]!.bytes.fill(0);
        expect(peer.surface!.session.currentPackage()).toBe(snapshot);
        const again = await readCommittedView(peer.surface!.session, token, {
          includeImageAssets: true,
        });
        if (!again.success) throw Error(again.error);
        expect(again.view.media![0]!.bytes).toEqual(galleryPng);
      }
    } finally {
      pair.destroy();
    }
  });

test('bounded asset preparation accepts a foreign byte realm and refuses spoofed or oversized views', async () => {
  const foreign = runInNewContext('new Uint8Array(bytes)', {
    bytes: [...galleryPng],
  }) as Uint8Array;
  expect(foreign instanceof Uint8Array).toBe(false);
  const id = sha256Bytes(galleryPng).replace(/^sha256:/, ''),
    path = `media/${id}.png`;
  const resolve = imageAssetBundleResolver([{ id, path, bytes: foreign, mimeType: 'image/png' }]);
  foreign.fill(0);
  expect((await resolve(path)).bytes).toEqual(galleryPng);
  const resources = await resolveInlineImageResources(
    [path],
    async () => ({
      bytes: runInNewContext('new Uint8Array(bytes)', { bytes: [...galleryPng] }),
      mime: 'image/png',
    }),
    galleryDecodePort
  );
  expect(resources.get(path)!.bytes).toEqual(galleryPng);
  for (const type of ['Uint8Array', 'Uint8ClampedArray', 'Int8Array']) {
    const bytes = runInNewContext(`new ${type}(bytes)`, {
      bytes: [...galleryPng],
    }) as Uint8Array;
    for (const key of ['length', 'byteLength', Symbol.iterator])
      Object.defineProperty(bytes, key, {
        get() {
          throw Error('Caller-owned byte access must not run.');
        },
      });
    const copy = imageAssetBundleResolver([{ id, path, bytes, mimeType: 'image/png' }]);
    expect((await copy(path)).bytes).toEqual(galleryPng);
  }
  for (const bytes of [
    new Uint16Array(4),
    new DataView(new ArrayBuffer(4)),
    { length: 4, byteLength: 4, [Symbol.toStringTag]: 'Uint8Array' },
  ])
    expect(() =>
      imageAssetBundleResolver([{ id, path, bytes: bytes as Uint8Array, mimeType: 'image/png' }])
    ).toThrow();
});
