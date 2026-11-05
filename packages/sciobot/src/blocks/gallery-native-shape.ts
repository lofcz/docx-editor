import type { JsonObject, JsonValue } from '../data/json';

export const NATIVE_GALLERY_VERSION = 3;
export const NATIVE_GALLERY_RESOURCE = Object.freeze({ version: 1, kind: 'embedded' });

/** This marker selects native ownership. It contains no bytes or address. */
export function nativeGalleryResource(value: JsonValue | undefined): boolean {
  return (
    !!value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    (value as JsonObject).version === 1 &&
    (value as JsonObject).kind === 'embedded' &&
    Object.keys(value).length === 2
  );
}

/** Optional fields are a read cache. Storage retains only the native resource marker. */
export function nativeGalleryImage(value: JsonValue): value is JsonObject {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const { resource, width, height, aspectRatio, caption, alt, title, ...extra } =
    value as JsonObject;
  return (
    Object.keys(extra).length === 0 &&
    nativeGalleryResource(resource) &&
    [width, height, aspectRatio].every(
      (size) =>
        size === undefined || (typeof size === 'number' && Number.isFinite(size) && size > 0)
    ) &&
    [caption, alt].every((text) => text === undefined || typeof text === 'string') &&
    (title === undefined || title === null || typeof title === 'string')
  );
}

export function nativeGalleryRecordShape(
  value: JsonObject,
  validId: (id: unknown) => boolean
): boolean {
  const attributes = value.attributes;
  if (!attributes || typeof attributes !== 'object' || Array.isArray(attributes)) return false;
  if (value.type === 'image_gallery_item')
    return (
      validId(value.galleryId) && value.galleryId !== value.id && nativeGalleryImage(attributes)
    );
  if (value.type !== 'image_gallery') return false;
  const { layout, align, images, ...extra } = attributes as JsonObject;
  return (
    Object.keys(extra).length === 0 &&
    (layout === 'row-fit' || layout === 'grid') &&
    (align === undefined || align === 'left' || align === 'center' || align === 'right') &&
    (images === undefined ||
      (Array.isArray(images) && images.length <= 100 && images.every(nativeGalleryImage)))
  );
}
