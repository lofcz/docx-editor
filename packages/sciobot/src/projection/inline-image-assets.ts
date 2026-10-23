import { forEachSemanticDrawing } from '@docx-editor.dev/core/layout';
import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import {
  findNode,
  resolveImageRelationship,
  sniffImageMime,
  validateRasterHeader,
  type OoxmlNode,
  type OoxmlPackage,
  type SupportedImageMime,
} from '@docx-editor.dev/core/store';
import type {
  MarkdownImageAsset,
  MarkdownImageOccurrence,
} from '@docx-editor.dev/docx-to-markdown';
import { decodeInlineImagePayload, INLINE_IMAGE } from '../blocks/inline-image-codec';
import { encodeBlockPayload } from '../blocks/payload-schema';
import type { IndexedBlock } from '../blocks/block-index';
const extensions: Readonly<Record<string, string>> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
};
function drawings(node: OoxmlNode): readonly OoxmlNode[] {
  return node.kind === 'drawing'
    ? [node]
    : node.kind === 'textValue'
      ? []
      : node.children.flatMap(drawings);
}
function pictureRelationship(node: OoxmlNode): readonly string[] {
  if (node.kind === 'textValue') return [];
  const id =
    node.localName === 'blip'
      ? node.attributes.find(
          (attr) =>
            attr.localName === 'embed' &&
            attr.namespaceUri ===
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        )?.value
      : undefined;
  return [...(id ? [id] : []), ...node.children.flatMap(pictureRelationship)];
}
/** Supply only managed pictures to the native renderer. Imported unregistered drawings stay omitted. */
export function managedInlineImageAssets(
  pkg: OoxmlPackage,
  layout: ExportSemanticLayout,
  blocks: readonly IndexedBlock[]
): readonly MarkdownImageAsset[] | undefined {
  const owners = new Map<
    string,
    { block: IndexedBlock; data: NonNullable<ReturnType<typeof decodeInlineImagePayload>> }
  >();
  for (const block of blocks) {
    if (block.payload.type !== INLINE_IMAGE) continue;
    const data = decodeInlineImagePayload(encodeBlockPayload(block.payload));
    const part = pkg.parts.get(block.partName),
      control = part && findNode(part, block.controlId);
    if (!data || !control) throw Error('The managed image control or provenance is invalid.');
    const native = drawings(control);
    if (native.length !== 1)
      throw Error('A managed inline image needs exactly one native drawing.');
    owners.set(block.partName + '\0' + native[0]!.id, { block, data });
  }
  if (!owners.size) return undefined;
  const assets = new Map<
    string,
    { data: Omit<MarkdownImageAsset, 'occurrences'>; occurrences: MarkdownImageOccurrence[] }
  >();
  forEachSemanticDrawing(layout, ({ drawing, page, story, rootStory }) => {
    const owner = owners.get(drawing.ownerPartName + '\0' + drawing.drawingNodeId);
    if (!owner || drawing.accessibility.hidden) return;
    const part = pkg.parts.get(owner.block.partName)!;
    const ids = pictureRelationship(findNode(part, drawing.drawingNodeId)!);
    if (ids.length !== 1)
      throw Error('The managed picture has no unique native media relationship.');
    const resolved = resolveImageRelationship(
      pkg.relationships.get(owner.block.partName),
      owner.block.partName,
      ids[0]!
    );
    if (resolved.mode !== 'internal')
      throw Error('A managed picture must use an embedded media part.');
    const bytes =
      pkg.partBytes.get(resolved.partName) ??
      pkg.partBytes.get(resolved.partName.replace(/^\//, ''));
    const mime = bytes && sniffImageMime(bytes),
      header = bytes && mime && validateRasterHeader(bytes, mime as SupportedImageMime);
    if (!bytes || !mime || !header || !extensions[mime])
      throw Error('The managed image media cannot be represented.');
    const id = owner.data.attributes.contentHash;
    let asset = assets.get(id);
    if (!asset) {
      const path = `media/${id}.${extensions[mime]}`;
      asset = {
        data: {
          id,
          path,
          url: path,
          mimeType: mime,
          bytes,
          byteLength: bytes.byteLength,
          pixelWidth: header.pixelWidth,
          pixelHeight: header.pixelHeight,
        },
        occurrences: [],
      };
      assets.set(id, asset);
    } else if (
      asset.data.bytes.length !== bytes.length ||
      asset.data.bytes.some((byte, index) => byte !== bytes[index])
    )
      throw Error('Managed image content identities conflict.');
    asset.occurrences.push({
      pageNumber: page.index + 1,
      story,
      rootStory,
      partName: drawing.ownerPartName,
      drawingNodeId: drawing.drawingNodeId,
      paragraphId: drawing.paragraphId,
      start: drawing.start,
      displayWidthPx: (drawing.width * 96) / 72,
      displayHeightPx: (drawing.height * 96) / 72,
      kind: drawing.kind === 'inlineDrawing' ? 'inline' : 'anchored',
      decorative: drawing.accessibility.decorative,
      alt: drawing.accessibility.label ?? '',
    });
  });
  return [...assets.values()].map(({ data, occurrences }) =>
    Object.freeze({ ...data, occurrences: Object.freeze(occurrences) })
  );
}
