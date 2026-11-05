import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  contentControlContentNodeOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { decodeBlockPayload as oldDecode } from '../../test/compat/payload-reader-v1';
import {
  GALLERY_IMAGES,
  galleryDocument,
  legacyGallery,
  openGallery,
} from '../../test/support/gallery';
import { prepareNativeDocument } from '../native/prepare-document';
import { buildBlockIndex } from './block-index';
import { decodeBlockPayload, encodeBlockPayload } from './payload-schema';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { withGalleryImageRecords } from './gallery-record-upgrade';

test('gallery bootstrap upgrades deterministically and preserves native paragraph owners', () => {
  const original = openGallery(legacyGallery());
  const oldBlock = buildBlockIndex(original).byId.get('gallery')!;
  const part = original.parts.get(oldBlock.partName)!;
  const oldControl = findNode(part, oldBlock.controlId)!;
  const oldParagraphs = storyParagraphs(contentControlContentNodeOf(oldControl)!).map(
    (paragraph) => paragraph.id
  );
  expect(oldParagraphs).toHaveLength(3);
  const result = withGalleryImageRecords(original);
  if (!result.ok) throw Error(result.reason);
  const gallery = buildBlockIndex(result.pkg).byId.get('gallery')!;
  expect(gallery.controlId).toBe(oldBlock.controlId);
  expect(gallery.payload).toEqual({ ...oldBlock.payload, version: 2 });
  const nextPart = result.pkg.parts.get(gallery.partName)!;
  const control = findNode(nextPart, gallery.controlId)!;
  expect(
    storyParagraphs(contentControlContentNodeOf(control)!).map((paragraph) => paragraph.id)
  ).toEqual(oldParagraphs);
  const content = contentControlContentNodeOf(control)!;
  expect(content.children.every((child) => child.kind === 'contentControl')).toBe(true);
  const items = [...buildBlockIndex(result.pkg).byId.values()].filter(
    (block) => block.parentId === 'gallery'
  );
  expect(new Set(items.map((item) => item.id)).size).toBe(3);
  expect(items.map((item) => item.payload.attributes)).toEqual(GALLERY_IMAGES);
  expect(items.every((item) => item.payload.galleryId === 'gallery')).toBe(true);
  expect(withGalleryImageRecords(result.pkg)).toEqual({ ok: true, pkg: result.pkg });
  const repeat = withGalleryImageRecords(original);
  expect(repeat.ok && semanticDigest(repeat.pkg.parts.values())).toEqual(
    semanticDigest(result.pkg.parts.values())
  );
});

test('a pinned old reader keeps gallery version 2 opaque and preserves Word fallback on reopen', () => {
  const pkg = openGallery(galleryDocument());
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') throw Error('Missing store');
  for (const entry of store.entries.values()) {
    const old = oldDecode(entry.raw);
    expect(old.status).toBe('unsupported-version');
    expect('raw' in old && old.raw).toBe(entry.raw);
  }
  const reopened = openGallery(writeOoxmlPackage(pkg));
  const reopenedStore = readPayloadStore(reopened);
  if (reopenedStore.status !== 'ready') throw Error('Missing reopened store');
  expect([...reopenedStore.entries.values()].map((entry) => entry.raw)).toEqual(
    [...store.entries.values()].map((entry) => entry.raw)
  );
  expect(semanticDigest(reopened.parts.values())).toEqual(semanticDigest(pkg.parts.values()));
  const main = reopened.parts.get(reopened.mainDocumentPart)!;
  expect(
    storyParagraphs(bodyStoryRoot(main)!).map((paragraph) => paragraphTextOf(main, paragraph.id))
  ).toEqual(['First image', 'Second image', 'Third image', 'Tail']);
});

test('gallery schema accepts only implemented version 2 shapes and keeps resource limits', () => {
  for (const value of [
    { id: 'a', version: 2, type: 'note', attributes: {} },
    {
      id: 'a',
      version: 2,
      type: 'image_gallery_item',
      galleryId: 'a',
      attributes: GALLERY_IMAGES[0],
    },
    { id: 'a', version: 2, type: 'image_gallery', attributes: { layout: 'unimplemented' } },
    {
      id: 'a',
      version: 2,
      type: 'image_gallery',
      attributes: { layout: 'grid', images: Array(101).fill(GALLERY_IMAGES[0]) },
    },
  ])
    expect(decodeBlockPayload(JSON.stringify(value)).status).toBe('unsupported-version');
  expect(
    decodeBlockPayload(
      encodeBlockPayload({ id: 'v1', version: 1, type: 'note', attributes: { retained: true } })
    ).status
  ).toBe('supported');
  expect(
    decodeBlockPayload(
      JSON.stringify({
        id: 'a',
        version: 2,
        type: 'image_gallery',
        attributes: { layout: 'grid' },
      }),
      { maxBytes: 8 }
    ).status
  ).toBe('invalid');
});

test('whole-array writes refuse and repeated document preparation does not upgrade live records', () => {
  const document = galleryDocument();
  const pkg = openGallery(document);
  const gallery = buildBlockIndex(pkg).byId.get('gallery')!;
  expect(
    withPayloadRecord(
      pkg,
      { ...gallery.payload, attributes: { layout: 'grid', images: [...GALLERY_IMAGES].reverse() } },
      'Gallery'
    )
  ).toMatchObject({ ok: false, reason: 'gallery-images-require-owned-record-command' });
  const prepared = prepareNativeDocument(document);
  expect(prepared.success && prepared.changed).toBe(false);
  expect(prepared.success && readOoxmlPackage(prepared.document).ok).toBe(true);
});
