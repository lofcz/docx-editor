import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  parentNodeOf,
  paragraphTextOf,
  semanticDigest,
  storyParagraphs,
  writeOoxmlPackage,
  type ImageDecodePort,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { protectedDocument } from '../../test/support/protected-document';
import { galleryDocument, openGallery } from '../../test/support/gallery';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { ungroupGallery } from './gallery-ungroup';
import { buildBlockIndex } from './block-index';
import { removeBlockSelection } from './remove';

const mount = (document: Uint8Array) =>
  createDocxEditor({ container: globalThis.document.createElement('div'), document });
const digest = (editor: DocxEditorInstance) =>
  semanticDigest(editor.surface!.session.currentPackage().parts.values());
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

async function images() {
  const editor = mount(galleryDocument());
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
function imageParagraphs(pkg: OoxmlPackage) {
  return [...buildBlockIndex(pkg).byId.values()]
    .filter((block) => block.payload.type === 'inline_image')
    .map((block) => {
      const parent = parentNodeOf(pkg.parts.get(block.partName)!, block.controlId);
      if (!parent || parent.kind !== 'paragraph') throw Error('Missing image paragraph');
      return parent.id;
    });
}

function imageRange(pkg: OoxmlPackage) {
  const images = imageParagraphs(pkg);
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  const owner = parentNodeOf(part, images[0]!)!;
  const first = owner.children.findIndex((node) => node.id === images[0]);
  const last = owner.children.findIndex((node) => node.id === images.at(-1));
  return owner.children.slice(first, last + 1).map((node) => node.id);
}

test('deleting the final paragraph of a table cell keeps one empty paragraph and exact undo', async () => {
  const input = await targetDocumentInput(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>'
  );
  const editor = mount(writeOoxmlPackage(input.pkg));
  try {
    const surface = editor.surface!;
    const paragraphId = surface.session
      .paragraphIds()
      .find((id) => paragraphTextOf(surface.session.part(), id) === 'Cell')!;
    const cellId = parentNodeOf(surface.session.part(), paragraphId)!.id;
    const before = digest(editor);
    const result = removeBlockSelection(editor, { nodeIds: [paragraphId] });
    if (!result.success) throw Error(result.error);
    const cell = findNode(surface.session.part(), cellId)!;
    expect(storyParagraphs(cell)).toHaveLength(1);
    expect(paragraphTextOf(surface.session.part(), storyParagraphs(cell)[0]!.id)).toBe('');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(before);
  } finally {
    editor.destroy();
  }
});

test('view mode and active header scopes refuse body range commands without changes', () => {
  const editor = mount(protectedDocument());
  try {
    const surface = editor.surface!;
    const nodeId = surface.session.paragraphIds()[0]!;
    const before = digest(editor);
    surface.setEditingMode('view');
    expect(removeBlockSelection(editor, { nodeIds: [nodeId] }).success).toBe(false);
    expect(digest(editor)).toEqual(before);
    surface.setEditingMode('edit');
    expect(surface.setActiveScope({ kind: 'headerFooter', rId: 'header' })).toBe(true);
    const headerReady = digest(editor);
    expect(removeBlockSelection(editor, { nodeIds: [nodeId] })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    expect(digest(editor)).toEqual(headerReady);
    expect(surface.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('a managed inline-image paragraph range keeps exact payload and resource undo', async () => {
  const editor = mount(await images());
  try {
    const surface = editor.surface!;
    const before = digest(editor);
    const selected = imageRange(surface.session.currentPackage());
    const result = removeBlockSelection(editor, { nodeIds: selected });
    if (!result.success) throw Error(result.error);
    expect(result.removedPayloadIds).toHaveLength(3);
    expect(imageParagraphs(surface.session.currentPackage())).toEqual([]);
    const removed = digest(editor);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(removed);
    expect(semanticDigest(openGallery(new Uint8Array(await editor.save())).parts.values())).toEqual(
      removed
    );
  } finally {
    editor.destroy();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`managed inline-image range deletion keeps shared resources, peer text, and history (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({
      document: await images(),
      clientIds,
      ownedAttributes: true,
    });
    try {
      const [alice, bob] = pair.editors;
      const beforeImages = imageParagraphs(alice!.surface!.session.currentPackage()).map((id) =>
        canonicalOoxmlFingerprint(findNode(alice!.surface!.session.part(), id)!)
      );
      pair.pause();
      const tail = bob!.surface!.session.paragraphIds().at(-1)!;
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: tail, offset: 4, text: ' peer' },
        ]).committed
      ).toBe(true);
      const result = removeBlockSelection(alice!, {
        nodeIds: imageRange(alice!.surface!.session.currentPackage()),
      });
      if (!result.success) throw Error(result.error);
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors)
        expect(imageParagraphs(editor.surface!.session.currentPackage())).toEqual([]);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(
          imageParagraphs(editor.surface!.session.currentPackage()).map((id) =>
            canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), id)!)
          )
        ).toEqual(beforeImages);
        expect(
          paragraphTextOf(
            editor.surface!.session.part(),
            editor.surface!.session.paragraphIds().at(-1)!
          )
        ).toBe('Tail peer');
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      for (const editor of pair.editors)
        expect(imageParagraphs(openGallery(new Uint8Array(await editor.save())))).toEqual([]);
    } finally {
      pair.destroy();
    }
  });
