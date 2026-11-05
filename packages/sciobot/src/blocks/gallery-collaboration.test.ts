import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  findNode,
  paragraphTextOf,
  semanticDigest,
  storyParagraphs,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  GALLERY_IMAGES,
  galleryDocument,
  legacyGallery,
  openGallery,
} from '../../test/support/gallery';
import { prepareNativeDocument } from '../native/prepare-document';
import { buildBlockIndex } from './block-index';
import { readGallery, updateGallery, type GalleryChange } from './gallery';

type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
const state = (editor: DocxEditorInstance) =>
  readGallery(editor.surface!.session.currentPackage(), 'gallery')!;
const captions = (editor: DocxEditorInstance) =>
  state(editor).images.map((image) => String(image.attributes.caption ?? image.attributes.url));
const order = (editor: DocxEditorInstance) => state(editor).images.map((image) => image.id);
const digest = (editor: DocxEditorInstance) =>
  semanticDigest(editor.surface!.session.currentPackage().parts.values());
const textIdentity = (editor: DocxEditorInstance, imageId: string) => {
  const image = state(editor).images.find((item) => item.id === imageId)!;
  const root = findNode(editor.surface!.session.part(), image.controlId)!;
  const values: string[] = [];
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'textValue') values.push(node.id);
    else node.children.forEach(visit);
  };
  visit(root);
  expect(values).toHaveLength(1);
  return values[0]!;
};
const change = (editor: DocxEditorInstance, change: GalleryChange) => {
  expect(updateGallery(editor, { blockId: 'gallery', change })).toMatchObject({ success: true });
};
const caption = (editor: DocxEditorInstance, imageId: string, text: string) => {
  const image = state(editor).images.find((item) => item.id === imageId)!;
  change(editor, {
    type: 'caption',
    imageId,
    caption: text,
    expectedCaption: String(image.attributes.caption ?? ''),
  });
};
const fallback = (editor: DocxEditorInstance) => {
  const part = editor.surface!.session.part();
  return storyParagraphs(bodyStoryRoot(part)!).map((paragraph) =>
    paragraphTextOf(part, paragraph.id)
  );
};

async function verify(pair: Pair, expected: readonly string[]) {
  await pair.settle();
  for (const editor of pair.editors) {
    expect(captions(editor)).toEqual([...expected]);
    expect(fallback(editor)).toEqual([...expected, 'Tail']);
    expect(buildBlockIndex(editor.surface!.session.currentPackage()).diagnostics).toEqual([]);
  }
  expect(digest(pair.editors[0]!)).toEqual(digest(pair.editors[1]!));
  const saved = await pair.savedDigests();
  expect(saved[0]).toEqual(saved[1]);
}

async function reopen(pair: Pair, expected: readonly string[]) {
  const bytes = new Uint8Array(await pair.editors[0]!.save());
  const saved = readGallery(openGallery(bytes), 'gallery')!;
  expect(
    saved.images.map((image) => String(image.attributes.caption ?? image.attributes.url))
  ).toEqual([...expected]);
  pair.flush();
  const reopened = await createTwoEditorFixture({
    document: bytes,
    roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
  });
  try {
    await verify(reopened, expected);
  } finally {
    reopened.destroy();
  }
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const) {
  test(`concurrent replacement of one gallery caption stays atomic through undo ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const id = order(alice)[0]!;
      const textId = textIdentity(alice, id);
      expect(textIdentity(bob, id)).toBe(textId);
      for (const document of pair.documents)
        expect(
          document.getMap<Y.Map<unknown>>('docx-package-nodes-v1').get(textId)?.has('atomicText')
        ).toBe(true);
      pair.pause();
      caption(alice, id, 'Alice replacement');
      caption(bob, id, 'Bob replacement');
      expect(textIdentity(alice, id)).toBe(textId);
      expect(textIdentity(bob, id)).toBe(textId);
      pair.reconnect();
      await pair.settle();
      const merged = captions(alice);
      expect(['Alice replacement', 'Bob replacement']).toContain(merged[0]);
      await verify(pair, merged);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await verify(pair, ['Bob replacement', 'Second image', 'Third image']);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const redone = captions(alice);
      expect(['Alice replacement', 'Bob replacement']).toContain(redone[0]);
      await verify(pair, redone);
      await reopen(pair, redone);
    } finally {
      pair.destroy();
    }
  });
  test(`separate gallery captions retain native owners through peer undo and room reopen ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const ids = order(alice);
      const controls = state(alice).images.map((image) => image.controlId);
      pair.pause();
      caption(alice, ids[0]!, 'Alice caption');
      caption(bob, ids[1]!, 'Bob caption');
      const peerOnly = digest(bob);
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await verify(pair, ['Alice caption', 'Bob caption', 'Third image']);
      expect(state(alice).images.map((image) => image.controlId)).toEqual(controls);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await verify(pair, ['First image', 'Bob caption', 'Third image']);
      expect(digest(alice)).toEqual(peerOnly);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await verify(pair, ['Alice caption', 'Bob caption', 'Third image']);
      await reopen(pair, ['Alice caption', 'Bob caption', 'Third image']);
    } finally {
      pair.destroy();
    }
  });

  test(`gallery reorder moves only identities while a peer changes the moved caption ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const ids = order(alice);
      pair.pause();
      change(alice, { type: 'move', imageId: ids[0]!, beforeImageId: null, expectedOrder: ids });
      caption(bob, ids[0]!, 'Peer first');
      const peerOnly = digest(bob);
      pair.reconnect();
      await verify(pair, ['Second image', 'Third image', 'Peer first']);
      expect(order(alice)).toEqual([ids[1], ids[2], ids[0]]);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await verify(pair, ['Peer first', 'Second image', 'Third image']);
      expect(digest(alice)).toEqual(peerOnly);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await verify(pair, ['Second image', 'Third image', 'Peer first']);
      await reopen(pair, ['Second image', 'Third image', 'Peer first']);
    } finally {
      pair.destroy();
    }
  });

  test(`gallery deletion wins a peer caption or movement without corrupting remaining images ${clientIds}`, async () => {
    for (const other of ['caption', 'move'] as const) {
      const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const ids = order(alice);
        pair.pause();
        change(alice, { type: 'remove', imageId: ids[0]! });
        if (other === 'caption') caption(bob, ids[0]!, 'Peer deleted caption');
        else
          change(bob, { type: 'move', imageId: ids[0]!, beforeImageId: null, expectedOrder: ids });
        pair.reconnect();
        await verify(pair, ['Second image', 'Third image']);
        expect(
          updateGallery(bob, {
            blockId: 'gallery',
            change: {
              type: 'caption',
              imageId: ids[0]!,
              caption: 'Stale draft',
              expectedCaption: 'First image',
            },
          })
        ).toMatchObject({ success: false, code: 'stale-revision' });
        await reopen(pair, ['Second image', 'Third image']);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(state(alice).images).toHaveLength(3);
        expect(digest(alice)).toEqual(digest(bob));
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await verify(pair, ['Second image', 'Third image']);
      } finally {
        pair.destroy();
      }
    }
  });

  test(`concurrent gallery additions keep duplicate addresses as distinct owners ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      pair.pause();
      change(alice, {
        type: 'append',
        images: [{ id: 'alice-image', attributes: { ...GALLERY_IMAGES[0], caption: 'Alice new' } }],
      });
      change(bob, {
        type: 'append',
        images: [{ id: 'bob-image', attributes: { ...GALLERY_IMAGES[0], caption: 'Bob new' } }],
      });
      pair.reconnect();
      await pair.settle();
      const merged = captions(alice);
      expect(merged).toHaveLength(5);
      expect(new Set(order(alice)).size).toBe(5);
      expect(merged).toContain('Alice new');
      expect(merged).toContain('Bob new');
      await verify(pair, merged);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await verify(
        pair,
        merged.filter((value) => value !== 'Alice new')
      );
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await verify(pair, merged);
      await reopen(pair, merged);
    } finally {
      pair.destroy();
    }
  });
}

test('last gallery image removal deletes the gallery in one shared history operation', async () => {
  const prepared = prepareNativeDocument(legacyGallery([GALLERY_IMAGES[0]!]));
  if (!prepared.success) throw Error(prepared.error);
  const pair = await createTwoEditorFixture({ document: prepared.document });
  try {
    const alice = pair.editors[0]!;
    const before = digest(alice);
    change(alice, { type: 'remove', imageId: order(alice)[0]! });
    await pair.settle();
    for (const editor of pair.editors) {
      expect(readGallery(editor.surface!.session.currentPackage(), 'gallery')).toBeNull();
      expect(fallback(editor)).toEqual(['Tail']);
      expect(buildBlockIndex(editor.surface!.session.currentPackage()).diagnostics).toEqual([]);
    }
    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    await verify(pair, ['First image']);
    expect(digest(alice)).toEqual(before);
    expect(alice.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    expect(fallback(alice)).toEqual(['Tail']);
    const digests = await pair.savedDigests();
    expect(digests[0]).toEqual(digests[1]);
  } finally {
    pair.destroy();
  }
});
test('successive gallery image removals delete the gallery and restore each membership through history', async () => {
  const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds: [41, 73] });
  try {
    const alice = pair.editors[0]!;
    for (let count = 3; count > 0; count--) {
      const image = readGallery(alice.surface!.session.currentPackage(), 'gallery')!.images[0]!;
      const result = updateGallery(alice, {
        blockId: 'gallery',
        change: { type: 'remove', imageId: image.id },
      });
      expect(result).toMatchObject({ success: true });
      await pair.settle();
      for (const editor of pair.editors)
        expect(
          readGallery(editor.surface!.session.currentPackage(), 'gallery')?.images.length ?? 0
        ).toBe(count - 1);
    }
    for (let count = 1; count <= 3; count++) {
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(
          readGallery(editor.surface!.session.currentPackage(), 'gallery')!.images.length
        ).toBe(count);
    }
  } finally {
    pair.destroy();
  }
});
