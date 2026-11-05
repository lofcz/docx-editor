import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  findNode,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { galleryDocument, openGallery } from '../../test/support/gallery';
import { buildBlockIndex } from './block-index';
import { ungroupGallery } from './gallery-ungroup';
import { readManagedImage, updateInlineImage, type InlineImageUpdate } from './inline-image-update';
import { removeBlock } from './remove';

const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (value) => value.charCodeAt(0)
);
const decodePort: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
async function documentWithImages() {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: galleryDocument(),
    imageDecodePort: decodePort,
  });
  try {
    const result = await ungroupGallery(editor, {
      blockId: 'gallery',
      decodePort,
      resolve: async () => ({ bytes: png, mime: 'image/png' }),
    });
    if (!result.success) throw Error(result.error);
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
  }
}
const descendants = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants)),
];
function imageOf(pkg: OoxmlPackage) {
  const block = [...buildBlockIndex(pkg).byId.values()].find(
    (block) => block.payload.type === 'inline_image'
  )!;
  const drawing = descendants(findNode(pkg.parts.get(block.partName)!, block.controlId)!).find(
    (node) => node.kind === 'drawing'
  )!;
  return readManagedImage(pkg, block.partName, drawing.id)!;
}
function state(pkg: OoxmlPackage) {
  const image = imageOf(pkg);
  const nodes = descendants(image.drawing);
  const native = (kind: string, name: string) => {
    const node = nodes.find((node) => node.kind === kind);
    return node?.kind === 'textValue'
      ? undefined
      : node?.attributes.find((attribute) => attribute.localName === name)?.value;
  };
  expect(contentControlPropertiesOf(image.control).lock).toBe('sdtContentLocked');
  expect(Number(native('drawingExtent', 'cx'))).toBe(
    Math.round(image.payload.attributes.widthPoints * 12700)
  );
  expect(Number(native('drawingExtent', 'cy'))).toBe(
    Math.round(image.payload.attributes.heightPoints * 12700)
  );
  expect(native('drawingDocPr', 'descr') ?? '').toBe(image.payload.attributes.alt);
  return image.payload.attributes;
}
function update(editor: DocxEditorInstance, change: InlineImageUpdate['change']) {
  const session = editor.surface!.session;
  const image = imageOf(session.currentPackage());
  return updateInlineImage(editor, {
    drawingNodeId: image.drawing.id,
    expectedRevision: session.packageRevision(),
    change,
  });
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`managed image size and alt remain native through peer edits, history, deletion, and reopen ${clientIds.join('/')}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await documentWithImages(),
      clientIds,
      editor: { imageDecodePort: decodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const original = state(alice!.surface!.session.currentPackage());
      pair.pause();
      expect(update(alice!, { type: 'size', widthPoints: 120, heightPoints: 60 }).success).toBe(
        true
      );
      expect(update(bob!, { type: 'alt', alt: 'Peer description' }).success).toBe(true);
      pair.reconnect(0);
      await pair.settle();
      for (const editor of pair.editors)
        expect(state(editor.surface!.session.currentPackage())).toMatchObject({
          widthPoints: 120,
          heightPoints: 60,
          alt: 'Peer description',
        });
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(state(editor.surface!.session.currentPackage())).toMatchObject({
          widthPoints: original.widthPoints,
          heightPoints: original.heightPoints,
          alt: 'Peer description',
        });
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(state(openGallery(new Uint8Array(await editor.save())))).toMatchObject({
          widthPoints: 120,
          heightPoints: 60,
          alt: 'Peer description',
        });
      pair.pause();
      expect(update(alice!, { type: 'alt', alt: 'Alice description' }).success).toBe(true);
      expect(update(bob!, { type: 'alt', alt: 'Bob description' }).success).toBe(true);
      pair.reconnect(1);
      await pair.settle();
      expect(state(alice!.surface!.session.currentPackage())).toEqual(
        state(bob!.surface!.session.currentPackage())
      );
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(state(editor.surface!.session.currentPackage()).alt).toBe('Bob description');
      const deleted = imageOf(bob!.surface!.session.currentPackage());
      pair.pause();
      expect(update(alice!, { type: 'size', widthPoints: 200, heightPoints: 100 }).success).toBe(
        true
      );
      expect(removeBlock(bob!, { controlId: deleted.control.id, mode: 'delete' }).success).toBe(
        true
      );
      pair.reconnect(0);
      await pair.settle();
      for (const editor of pair.editors)
        expect(
          buildBlockIndex(editor.surface!.session.currentPackage()).byId.has(deleted.block.id)
        ).toBe(false);
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  });
