import type { BlockPayload } from '../blocks/payload-schema';
import type { JsonObject, JsonValue } from '../data/json';
import type { SerializedBlockTag } from './block-tags';
import { validGalleryAssetPayload } from './gallery-assets';

/** The application's gallery grammar: `layout`, optional `align`, and an `images` list. */
export const GALLERY_LAYOUTS: readonly string[] = ['row-fit', 'grid'];
export const GALLERY_ALIGNS: readonly string[] = ['left', 'center', 'right'];
const MAX_IMAGES = 100;
const MAX_TEXT = 2048;

const finite = (value: JsonValue | undefined) =>
  value === undefined || (typeof value === 'number' && Number.isFinite(value) && value > 0);

/**
 * One gallery image: an absolute http(s) URL and an optional caption and size. Other keys (such
 * as attribution) stay as data; the renderer never interprets them as markup.
 */
export function galleryImage(value: JsonValue): value is JsonObject {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const image = value as JsonObject;
  if (typeof image.url !== 'string' || image.url.length > MAX_TEXT) return false;
  try {
    if (!['http:', 'https:'].includes(new URL(image.url).protocol)) return false;
  } catch {
    return false;
  }
  return (
    (image.caption === undefined ||
      (typeof image.caption === 'string' && image.caption.length <= MAX_TEXT)) &&
    finite(image.width) &&
    finite(image.height) &&
    finite(image.aspectRatio)
  );
}

export function validGallery(payload: BlockPayload): boolean {
  const { layout, align, images, ...rest } = payload.attributes;
  return (
    Object.keys(rest).length === 0 &&
    typeof layout === 'string' &&
    GALLERY_LAYOUTS.includes(layout) &&
    (align === undefined || (typeof align === 'string' && GALLERY_ALIGNS.includes(align))) &&
    Array.isArray(images) &&
    images.length > 0 &&
    images.length <= MAX_IMAGES &&
    images.every(galleryImage)
  );
}

/** `images` is a JSX expression holding a JSON literal, as the application's Plate rules write it. */
export function serializeGallery(payload: BlockPayload): SerializedBlockTag {
  if (!validGallery(payload) && !validGalleryAssetPayload(payload))
    throw new TypeError('invalid-gallery-payload');
  const { layout, align, images, resourceVersion } = payload.attributes;
  const text =
    `<image_gallery layout="${layout as string}"` +
    (resourceVersion === undefined ? '' : ` resourceVersion="1"`) +
    (align === undefined ? '' : ` align="${align as string}"`) +
    ` images={${JSON.stringify(images)}} />`;
  return {
    tag: 'image_gallery',
    text,
    properties: [
      'layout',
      ...(resourceVersion === undefined ? [] : ['resourceVersion']),
      ...(align === undefined ? [] : ['align']),
      'images',
    ].map((property) => ({
      name: property,
      property,
      path: [property],
      start: 0,
      end: text.length,
      valueStart: null,
      valueEnd: null,
    })),
  };
}

/** Read `images` from a JSON string or an inert expression attribute; anything else is null. */
export function galleryImages(value: JsonValue | undefined): JsonValue | null {
  const source =
    typeof value === 'string'
      ? value
      : value &&
          typeof value === 'object' &&
          !Array.isArray(value) &&
          (value as JsonObject).type === 'mdxJsxAttributeValueExpression' &&
          typeof (value as JsonObject).value === 'string'
        ? ((value as JsonObject).value as string)
        : null;
  if (source === null || source.length > MAX_IMAGES * MAX_TEXT * 2) return null;
  try {
    return JSON.parse(source) as JsonValue;
  } catch {
    return null;
  }
}
