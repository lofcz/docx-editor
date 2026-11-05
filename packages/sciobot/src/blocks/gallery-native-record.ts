import {
  contentControlContentNodeOf,
  directParagraphProperties,
  paragraphTextOf,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { JsonObject } from '../data/json';
import { nativeImageFields, nativeImageResource } from './native-image-fields';
import { NATIVE_GALLERY_RESOURCE } from './gallery-native-shape';
import { isPlainGalleryFallback } from './gallery-record-build';

/** A standalone native picture paragraph can move without changing its drawing or run identities. */
export function nativeGalleryPicture(pkg: OoxmlPackage, part: OoxmlPart, paragraph: OoxmlNode) {
  if (paragraph.kind !== 'paragraph') return null;
  const pictures: OoxmlNode[] = [];
  const pending: OoxmlNode[] = [paragraph];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'drawing') {
      pictures.push(node);
      continue;
    }
    if (node.kind === 'textValue') {
      if (node.value.length) return null;
      continue;
    }
    if (['paragraphProperties', 'runProperties', 'contentControlProperties'].includes(node.kind))
      continue;
    if (
      !['paragraph', 'run', 'text', 'contentControl', 'contentControlContent'].includes(node.kind)
    )
      return null;
    for (const child of node.children) pending.push(child);
  }
  if (pictures.length !== 1) return null;
  // Only an inline picture is eligible. Charts, anchored shapes, and text boxes keep native chrome.
  const drawing = pictures[0]!;
  const inline = drawing.kind !== 'textValue' && drawing.children[0];
  if (
    drawing.kind === 'textValue' ||
    drawing.children.length !== 1 ||
    !inline ||
    inline.kind === 'textValue' ||
    inline.localName !== 'inline' ||
    inline.namespaceUri !== 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
  )
    return null;
  const graphic = inline.children.find(
    (node) => node.kind !== 'textValue' && node.localName === 'graphic'
  );
  const data =
    graphic &&
    graphic.kind !== 'textValue' &&
    graphic.children.find((node) => node.kind !== 'textValue' && node.localName === 'graphicData');
  if (
    !data ||
    data.kind === 'textValue' ||
    data.children.length !== 1 ||
    data.children[0]?.kind === 'textValue' ||
    data.children[0]?.localName !== 'pic' ||
    data.children[0]?.namespaceUri !== 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  )
    return null;
  const fields = nativeImageFields(pictures[0]!);
  const resource = nativeImageResource(pkg, part.name, pictures[0]!);
  if (
    !fields ||
    !resource ||
    (!pkg.partBytes.has(resource.partName) && !pkg.partBytes.has(resource.partName.slice(1)))
  )
    return null;
  return { paragraph, fields, resource };
}

export function hasNativeGalleryCaptionStyle(part: OoxmlPart, node: OoxmlNode): boolean {
  if (node.kind !== 'paragraph') return false;
  const style = directParagraphProperties(part, node.id).find(
    (property) => property.localName === 'pStyle'
  );
  return (style?.attributes?.val ?? style?.attributes?.['w:val']) === 'Caption';
}

export function isNativeGalleryCaption(part: OoxmlPart, node: OoxmlNode): boolean {
  return hasNativeGalleryCaptionStyle(part, node) && isPlainGalleryFallback(node);
}

/** Read one native image owner. Child order identifies its picture and optional caption. */
export function readNativeGalleryImage(pkg: OoxmlPackage, part: OoxmlPart, control: OoxmlNode) {
  const content = contentControlContentNodeOf(control);
  if (!content || content.children.length < 1 || content.children.length > 2) return null;
  const picture = nativeGalleryPicture(pkg, part, content.children[0]!);
  if (!picture) return null;
  const caption = content.children[1];
  if (caption && !isNativeGalleryCaption(part, caption)) return null;
  const { widthPoints, heightPoints, alt, title } = picture.fields.attributes;
  return {
    ...picture,
    caption: caption?.kind === 'paragraph' ? caption : null,
    attributes: {
      resource: NATIVE_GALLERY_RESOURCE,
      width: widthPoints / 0.75,
      height: heightPoints / 0.75,
      aspectRatio: widthPoints / heightPoints,
      caption: caption ? (paragraphTextOf(part, caption.id) ?? '') : '',
      alt,
      title,
    } satisfies JsonObject,
  };
}
