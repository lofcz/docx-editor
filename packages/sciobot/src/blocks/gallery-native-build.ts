import {
  bodyStoryRoot,
  contentControlsIn,
  createNodeIdAllocator,
  escapeXml,
  readOoxmlPart,
  usedParaIds,
  w14RootPrefix,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { embedInlineImage } from '../native/inline-image-media';
import { inlinePicture } from '../native/inline-picture';
import type { ResolvedInlineImage } from '../native/inline-image-resources';
import { blockTag, withPayloadRecord } from './payload-opc-codec';
import { createInlineImagePayload } from './inline-image-codec';
import { withFreshIds } from './fresh-ids';
import { nativeGalleryControl } from './gallery-group';
import { NATIVE_GALLERY_RESOURCE } from './gallery-native-shape';

/** Create native descendants from a prepared resource. Metadata contains provenance, never image bytes. */
export function buildNativeGalleryImage(
  pkg: OoxmlPackage,
  partName: string,
  input: {
    galleryId: string;
    imageId: string;
    inlineId: string;
    reference: string;
    resource: ResolvedInlineImage;
    caption?: string;
    alt?: string;
    title?: string | null;
  }
) {
  const embedded = embedInlineImage(pkg, partName, input.inlineId, input.resource);
  let next = embedded.pkg;
  const part = next.parts.get(partName)!;
  const caption = (input.caption ?? '').replace(/\r\n|\r/g, '\n');
  const alt = input.alt ?? caption;
  const drawing = inlinePicture({
    docPrId: embedded.docPrId,
    relationshipId: embedded.relationshipId,
    widthPoints: input.resource.widthPoints,
    heightPoints: input.resource.heightPoints,
    title: input.title ?? null,
    alt,
  });
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body><w:p><w:sdt><w:sdtPr><w:tag w:val="${escapeXml(blockTag(input.inlineId))}"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>\uFFFC</w:t></w:r></w:sdtContent></w:sdt></w:p><w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r>${caption
      .split('\n')
      .map((line) => `<w:t xml:space="preserve">${escapeXml(line)}</w:t>`)
      .join('<w:br/>')}</w:r></w:p></w:body></w:document>`,
    { name: '/sciobot-gallery-native-image.xml', contentType: part.contentType }
  );
  const body = parsed.ok && bodyStoryRoot(parsed.part);
  if (!body || body.kind === 'textValue' || body.children.length !== 2)
    throw Error('The native image template is invalid.');
  const replace = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'text'
      ? drawing
      : node.kind === 'textValue'
        ? node
        : ({ ...node, children: node.children.map(replace) } as OoxmlNode);
  const allocate = createNodeIdAllocator(part, 'new');
  const identity = {
    prefix: w14RootPrefix(part.root),
    used: new Set(usedParaIds(part.root)),
    seed: `gallery-image:${input.imageId}`,
  };
  const paragraphs = [
    withFreshIds(replace(body.children[0]!), allocate, identity),
    withFreshIds(body.children[1]!, allocate, identity),
  ];
  const payload = createInlineImagePayload(input.inlineId, {
    reference: input.reference,
    alt,
    title: input.title ?? null,
    contentHash: input.resource.contentHash,
    mediaPartName: embedded.partName.startsWith('/') ? embedded.partName : '/' + embedded.partName,
    widthPoints: input.resource.widthPoints,
    heightPoints: input.resource.heightPoints,
  });
  const stored = withPayloadRecord(next, payload, alt.slice(0, 256) || 'Image');
  if (!stored.ok) throw Error(stored.reason);
  next = stored.pkg;
  const owner = withPayloadRecord(
    next,
    {
      id: input.imageId,
      version: 3,
      type: 'image_gallery_item',
      galleryId: input.galleryId,
      attributes: { resource: NATIVE_GALLERY_RESOURCE },
    },
    'Image'
  );
  if (!owner.ok) throw Error(owner.reason);
  const inlineControlId = contentControlsIn(paragraphs[0]!)[0]?.node.id;
  if (!inlineControlId) throw Error('The native image control is missing.');
  return {
    pkg: owner.pkg,
    control: nativeGalleryControl(part, input.imageId, paragraphs, allocate),
    inlineControlId,
  };
}
