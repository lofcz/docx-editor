import {
  bodyStoryRoot,
  createNodeIdAllocator,
  escapeXml,
  readOoxmlPart,
  usedParaIds,
  w14RootPrefix,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { JsonObject } from '../data/json';
import { withFreshIds } from './fresh-ids';
import { blockTag } from './payload-opc-codec';

export function galleryImageText(image: JsonObject): string {
  return (image.caption as string | undefined) || (image.url as string);
}

/** One stable native owner per image, including duplicate addresses. */
export function galleryImageControl(
  part: OoxmlPart,
  imageId: string,
  image: JsonObject,
  paragraph?: OoxmlNode,
  allocate = createNodeIdAllocator(part, 'new')
): OoxmlElement {
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>` +
      `<w:sdt><w:sdtPr><w:tag w:val="${escapeXml(blockTag(imageId))}"/></w:sdtPr>` +
      `<w:sdtContent><w:p><w:r><w:t xml:space="preserve">${escapeXml(galleryImageText(image))}</w:t></w:r></w:p>` +
      '</w:sdtContent></w:sdt></w:body></w:document>',
    { name: '/sciobot-gallery-image.xml', contentType: part.contentType }
  );
  const body = parsed.ok && bodyStoryRoot(parsed.part);
  const template = body && body.kind !== 'textValue' && body.children[0];
  if (!template || template.kind !== 'contentControl')
    throw Error('Invalid gallery image fallback.');
  const control = withFreshIds(template, allocate, {
    prefix: w14RootPrefix(part.root),
    used: new Set(usedParaIds(part.root)),
    seed: `scio-gallery-image:${imageId}`,
  }) as OoxmlElement;
  if (!paragraph) return control;
  return {
    ...control,
    children: control.children.map((node) =>
      node.kind === 'contentControlContent'
        ? ({ ...node, children: [paragraph] } as OoxmlNode)
        : node
    ),
  } as OoxmlElement;
}
