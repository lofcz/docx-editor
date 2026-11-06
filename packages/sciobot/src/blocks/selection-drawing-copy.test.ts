import { expect, test } from 'bun:test';
import {
  findNode,
  parentNodeOf,
  paragraphTextOf,
  semanticDigest,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { openGallery } from '../../test/support/gallery';
import {
  galleryDescendants,
  galleryDecodePort,
  nativeGalleryDocument,
  nativeGalleryDrawings,
} from '../../test/support/native-gallery';
import { duplicateBlockSelection } from './duplicate';
import { isNativeGalleryCaption } from './gallery-native-record';

function imageRange(pkg: OoxmlPackage) {
  const drawings = nativeGalleryDrawings(pkg);
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  const paragraphs = drawings.map(({ block }) => parentNodeOf(part, block.controlId)!);
  const owner = parentNodeOf(part, paragraphs[0]!.id)!;
  const first = owner.children.findIndex((node) => node.id === paragraphs[0]!.id);
  let last = owner.children.findIndex((node) => node.id === paragraphs.at(-1)!.id) + 1;
  if (owner.children[last] && isNativeGalleryCaption(part, owner.children[last]!)) last++;
  return owner.children.slice(first, last).map((node) => node.id);
}
function drawingIds(pkg: OoxmlPackage) {
  const all = [...pkg.parts.values()].flatMap((part) => galleryDescendants(part.root));
  return {
    nodes: all.filter((node) => node.kind === 'drawing').map((node) => node.id),
    word: all
      .filter((node) => node.kind === 'drawingDocPr')
      .map((node) => node.attributes.find((attribute) => attribute.localName === 'id')!.value),
  };
}
function media(pkg: OoxmlPackage) {
  return [...pkg.partBytes]
    .filter(([name]) => name.includes('/media/'))
    .map(([name, bytes]) => [name, Array.from(bytes)]);
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`ordinary image ranges copy resources and keep peer text independent through history (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({
      document: await nativeGalleryDocument(),
      clientIds,
      ownedAttributes: true,
      editor: { imageDecodePort: galleryDecodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const surface = alice!.surface!,
        before = surface.session.currentPackage();
      const ids = imageRange(before),
        beforeMedia = media(before);
      const captionId = ids.find((id) =>
        isNativeGalleryCaption(surface.session.part(), findNode(surface.session.part(), id)!)
      )!;
      const caption = paragraphTextOf(surface.session.part(), captionId)!;
      pair.pause();
      let serial = 0;
      const copied = duplicateBlockSelection(alice!, {
        nodeIds: ids,
        allocateBlockId: (type) => `range-${type}-${++serial}`,
      });
      if (!copied.success) throw Error(copied.error);
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: captionId, offset: caption.length, text: ' peer' },
        ]).committed
      ).toBe(true);
      const peerOnly = semanticDigest(bob!.surface!.session.currentPackage().parts.values());
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors) {
        const part = editor.surface!.session.part(),
          pkg = editor.surface!.session.currentPackage();
        expect(paragraphTextOf(part, captionId)).toBe(`${caption} peer`);
        expect(copied.nodeIds.map((id) => paragraphTextOf(part, id))).toContain(caption);
        const drawing = drawingIds(pkg);
        expect(drawing.nodes).toHaveLength(6);
        expect(new Set(drawing.nodes).size).toBe(6);
        expect(new Set(drawing.word).size).toBe(6);
        expect(media(pkg)).toEqual(beforeMedia);
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(semanticDigest(editor.surface!.session.currentPackage().parts.values())).toEqual(
          peerOnly
        );
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const reopened = openGallery(new Uint8Array(await editor.save())),
          drawing = drawingIds(reopened);
        expect(drawing.nodes).toHaveLength(6);
        expect(new Set(drawing.word).size).toBe(6);
        expect(media(reopened)).toEqual(beforeMedia);
        expect(semanticDigest(reopened.parts.values())).toEqual(
          semanticDigest(editor.surface!.session.currentPackage().parts.values())
        );
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
