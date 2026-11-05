import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  findNode,
  parentNodeOf,
  type ImageDecodePort,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { galleryDocument } from '../../test/support/gallery';
import { buildBlockIndex } from './block-index';
import { ungroupGallery } from './gallery-ungroup';
import { readManagedImage, updateInlineImage } from './inline-image-update';
import { replaceInlineImage } from './inline-image-replace';
import { nativeImageResource } from './native-image-fields';
import { removeBlock } from './remove';

const bytes = (base64: string) => Uint8Array.from(atob(base64), (value) => value.charCodeAt(0));
const png = bytes(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
);
const gif = bytes('R0lGODlhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==');
const decodePort: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
async function fixture() {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: galleryDocument(),
    imageDecodePort: decodePort,
  });
  try {
    expect(
      (
        await ungroupGallery(editor, {
          blockId: 'gallery',
          decodePort,
          resolve: async () => ({ bytes: png, mime: 'image/png' }),
        })
      ).success
    ).toBe(true);
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
  }
}
const descendants = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants)),
];
function imageOf(editor: DocxEditorInstance, id?: string) {
  const pkg = editor.surface!.session.currentPackage();
  const block = [...buildBlockIndex(pkg).byId.values()].find(
    (item) => item.payload.type === 'inline_image' && (!id || item.id === id)
  )!;
  const control = findNode(pkg.parts.get(block.partName)!, block.controlId)!;
  const drawing = descendants(control).find((node) => node.kind === 'drawing')!;
  const image = readManagedImage(pkg, block.partName, drawing.id);
  expect(image).not.toBeNull();
  return image!;
}
async function state(editor: DocxEditorInstance, id?: string) {
  const image = imageOf(editor, id);
  const pkg = editor.surface!.session.currentPackage();
  const resource = nativeImageResource(pkg, image.part.name, image.control)!;
  expect(resource.partName).toBe(image.payload.attributes.mediaPartName);
  expect(contentControlPropertiesOf(image.control).lock).toBe('sdtContentLocked');
  const data =
    pkg.partBytes.get(resource.partName) ?? pkg.partBytes.get(resource.partName.slice(1));
  expect(data).toBeDefined();
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', new Uint8Array(data!)));
  expect(image.payload.attributes.contentHash).toBe(
    [...digest].map((value) => value.toString(16).padStart(2, '0')).join('')
  );
  const metadata = descendants(image.drawing).find((node) => node.kind === 'drawingDocPr')!;
  expect(
    metadata.attributes.find((attribute) => attribute.localName === 'descr')?.value ?? ''
  ).toBe(image.payload.attributes.alt);
  return image.payload.attributes;
}
function replace(editor: DocxEditorInstance, reference: string, data = gif) {
  const image = imageOf(editor);
  return replaceInlineImage(editor, {
    drawingNodeId: image.drawing.id,
    expectedRevision: editor.surface!.session.packageRevision(),
    reference,
    decodePort,
    resolve: async () => ({ bytes: data, mime: data === gif ? 'image/gif' : 'image/png' }),
  });
}

function selectImage(editor: DocxEditorInstance) {
  const image = imageOf(editor);
  let paragraph = parentNodeOf(image.part, image.drawing.id);
  while (paragraph && paragraph.kind !== 'paragraph')
    paragraph = parentNodeOf(image.part, paragraph.id);
  expect(paragraph?.kind).toBe('paragraph');
  expect(editor.surface!.selectDrawing(image.drawing.id, paragraph!.id)).toBe(true);
  expect(editor.getSelectedImage()?.id).toBe(image.drawing.id);
}

test('the same source reference can replace bytes without changing image geometry or description', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: await fixture(),
    imageDecodePort: decodePort,
  });
  try {
    const original = await state(editor);
    expect((await replace(editor, original.reference, gif)).success).toBe(true);
    const replaced = await state(editor);
    expect(replaced).toMatchObject({
      reference: original.reference,
      alt: original.alt,
      widthPoints: original.widthPoints,
      heightPoints: original.heightPoints,
    });
    expect(replaced.contentHash).not.toBe(original.contentHash);
    expect((await replace(editor, original.reference, gif)).success).toBe(true);
    expect(await state(editor)).toEqual(replaced);
  } finally {
    editor.destroy();
  }
});

for (const mode of ['view', 'suggest', 'unsupported-scope'] as const)
  test(`manual replacement refuses ${mode} before preparing resources`, async () => {
    const pair = await createTwoEditorFixture({
      document: await fixture(),
      ownedAttributes: mode !== 'unsupported-scope',
    });
    try {
      const editor = pair.editors[0]!;
      const image = imageOf(editor);
      if (mode !== 'unsupported-scope') editor.surface!.setEditingMode(mode);
      const before = editor.surface!.session.currentPackage();
      const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      let loaded = false;
      expect(
        await replaceInlineImage(editor, {
          drawingNodeId: image.drawing.id,
          expectedRevision: editor.surface!.session.packageRevision(),
          reference: 'asset:refused',
          decodePort,
          resolve: async () => {
            loaded = true;
            return { bytes: gif, mime: 'image/gif' };
          },
        })
      ).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(loaded).toBe(false);
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
    } finally {
      pair.destroy();
    }
  });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`manual replacement keeps peer metadata, history, native resources, and reopen ${clientIds.join('/')}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await fixture(),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: decodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const original = await state(alice!);
      const id = imageOf(alice!).block.id;
      pair.pause();
      selectImage(alice!);
      const field = document.createElement('input');
      alice!.surface!.hostElement().parentElement!.append(field);
      field.focus();
      expect((await replace(alice!, 'asset:replacement')).success).toBe(true);
      expect(alice!.getSelectedImage()?.id).toBe(imageOf(alice!).drawing.id);
      expect(document.activeElement).toBe(field);
      field.remove();
      selectImage(bob!);
      const peerField = document.createElement('input');
      bob!.surface!.hostElement().parentElement!.append(peerField);
      peerField.focus();
      let peer = imageOf(bob!);
      expect(
        updateInlineImage(bob!, {
          drawingNodeId: peer.drawing.id,
          expectedRevision: bob!.surface!.session.packageRevision(),
          change: { type: 'alt', alt: 'Peer description' },
        }).success
      ).toBe(true);
      expect(bob!.getSelectedImage()?.id).toBe(imageOf(bob!).drawing.id);
      expect(document.activeElement).toBe(peerField);
      peer = imageOf(bob!);
      expect(
        updateInlineImage(bob!, {
          drawingNodeId: peer.drawing.id,
          expectedRevision: bob!.surface!.session.packageRevision(),
          change: { type: 'size', widthPoints: 120, heightPoints: 60 },
        }).success
      ).toBe(true);
      expect(bob!.getSelectedImage()?.id).toBe(imageOf(bob!).drawing.id);
      expect(document.activeElement).toBe(peerField);
      peerField.remove();
      pair.reconnect(0);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await state(editor, id)).toMatchObject({
          reference: 'asset:replacement',
          alt: 'Peer description',
          widthPoints: 120,
          heightPoints: 60,
        });
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await state(editor, id)).toMatchObject({
          reference: original.reference,
          alt: 'Peer description',
          widthPoints: 120,
          heightPoints: 60,
        });
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const expected = await state(alice!, id);
      expect(await state(bob!, id)).toEqual(expected);
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: new Uint8Array(await alice!.save()),
        imageDecodePort: decodePort,
      });
      try {
        expect(await state(reopened, id)).toEqual(expected);
      } finally {
        reopened.destroy();
      }
    } finally {
      pair.destroy();
    }
  });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`same-byte replacement and another resource retain one native winner ${clientIds.join('/')}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await fixture(),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: decodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      pair.pause();
      expect((await replace(alice!, 'asset:new-bitmap', gif)).success).toBe(true);
      expect((await replace(bob!, 'asset:same-bitmap-new-reference', png)).success).toBe(true);
      pair.reconnect(0);
      await pair.settle();
      const winner = await state(alice!);
      expect(await state(bob!)).toEqual(winner);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(await state(alice!)).toEqual(await state(bob!));
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(await state(alice!)).toEqual(await state(bob!));
    } finally {
      pair.destroy();
    }
  });

test('replacement preparation refuses peer deletion and document replacement without writes', async () => {
  const pair = await createTwoEditorFixture({
    document: await fixture(),
    ownedAttributes: true,
    editor: { imageDecodePort: decodePort },
  });
  try {
    const [alice, bob] = pair.editors;
    const image = imageOf(alice!);
    selectImage(alice!);
    let release!: () => void;
    const pending = replaceInlineImage(alice!, {
      drawingNodeId: image.drawing.id,
      expectedRevision: alice!.surface!.session.packageRevision(),
      reference: 'asset:delayed',
      decodePort,
      resolve: async () => {
        await new Promise<void>((resolve) => {
          release = resolve;
        });
        return { bytes: gif, mime: 'image/gif' };
      },
    });
    const peer = imageOf(bob!);
    expect(removeBlock(bob!, { controlId: peer.control.id, mode: 'delete' }).success).toBe(true);
    await pair.settle();
    const before = alice!.surface!.session.currentPackage();
    const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    release();
    expect(await pending).toMatchObject({ success: false, code: 'stale-revision' });
    expect(alice!.surface!.session.currentPackage()).toBe(before);
    expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
    expect(alice!.getSelectedImage()).toBeNull();
    const remaining = imageOf(alice!);
    selectImage(alice!);
    const replacing = replaceInlineImage(alice!, {
      drawingNodeId: remaining.drawing.id,
      expectedRevision: alice!.surface!.session.packageRevision(),
      reference: 'asset:other-document',
      decodePort,
      resolve: async () => {
        await new Promise<void>((resolve) => {
          release = resolve;
        });
        return { bytes: gif, mime: 'image/gif' };
      },
    });
    await alice!.load(new Uint8Array(await alice!.save()));
    const owner = alice!.surface;
    const reopened = owner!.session.currentPackage();
    release();
    expect(await replacing).toMatchObject({ success: false, code: 'stale-revision' });
    expect(alice!.surface).toBe(owner);
    expect(owner!.session.currentPackage()).toBe(reopened);
    expect(alice!.getSelectedImage()).toBeNull();
  } finally {
    pair.destroy();
  }
});

for (const protection of ['contentLocked', 'sdtContentLocked', 'sdtLocked', 'dataBinding'])
  test(`replacement preserves ancestor ${protection} before resource loading`, async () => {
    const files = unzipSync(await fixture());
    const property =
      protection === 'dataBinding'
        ? '<w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000001}" w:xpath="/value"/>'
        : `<w:lock w:val="${protection}"/>`;
    files['word/document.xml'] = strToU8(
      strFromU8(files['word/document.xml']!)
        .replace('<w:body>', `<w:body><w:sdt><w:sdtPr>${property}</w:sdtPr><w:sdtContent>`)
        .replace('</w:body>', '</w:sdtContent></w:sdt></w:body>')
    );
    const pair = await createTwoEditorFixture({ document: zipSync(files), ownedAttributes: true });
    try {
      const editor = pair.editors[0]!;
      const image = imageOf(editor);
      const pkg = editor.surface!.session.currentPackage();
      const shared = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
      let loaded = false;
      expect(
        await replaceInlineImage(editor, {
          drawingNodeId: image.drawing.id,
          expectedRevision: editor.surface!.session.packageRevision(),
          reference: 'asset:refused',
          decodePort,
          resolve: async () => {
            loaded = true;
            return { bytes: gif, mime: 'image/gif' };
          },
        })
      ).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(loaded).toBe(false);
      expect(editor.surface!.session.currentPackage()).toBe(pkg);
      expect(pair.documents.map((doc) => Y.encodeStateAsUpdate(doc))).toEqual(shared);
    } finally {
      pair.destroy();
    }
  });
