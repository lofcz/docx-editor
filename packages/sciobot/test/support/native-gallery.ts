import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  findNode,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { galleryDocument } from './gallery';
import { ungroupGallery } from '../../src/blocks/gallery-ungroup';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { readGallery, updateGallery } from '../../src/blocks/gallery';

export const galleryPng = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (value) => value.charCodeAt(0)
);
export const galleryDecodePort: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
export async function nativeGalleryDocument(caption?: string) {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: galleryDocument(),
    imageDecodePort: galleryDecodePort,
  });
  try {
    if (caption !== undefined) {
      const image = readGallery(editor.surface!.session.currentPackage(), 'gallery')!.images[0]!;
      const changed = updateGallery(editor, {
        blockId: 'gallery',
        change: {
          type: 'caption',
          imageId: image.id,
          expectedCaption: String(image.attributes.caption),
          caption,
        },
      });
      if (!changed.success) throw Error(changed.error);
    }
    const result = await ungroupGallery(editor, {
      blockId: 'gallery',
      decodePort: galleryDecodePort,
      resolve: async () => ({ bytes: galleryPng, mime: 'image/png' }),
    });
    if (!result.success) throw Error(result.error);
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
  }
}
export const galleryDescendants = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(galleryDescendants)),
];
export function nativeGalleryDrawings(pkg: OoxmlPackage) {
  return [...buildBlockIndex(pkg).byId.values()]
    .filter((block) => block.payload.type === 'inline_image')
    .map((block) => ({
      block,
      drawing: galleryDescendants(findNode(pkg.parts.get(block.partName)!, block.controlId)!).find(
        (node) => node.kind === 'drawing'
      )!,
    }));
}
