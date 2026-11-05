import type { JsonObject, JsonValue } from '../data/json';
import { galleryImage } from '../projection/gallery';

export const GALLERY_IMAGE_RECORD = 'image_gallery_item';
export const GALLERY_RECORD_VERSION = 2;

/** Version 2 only names the implemented gallery and image-owner records. */
export function galleryRecordShape(value: JsonObject, validId: (id: unknown) => boolean): boolean {
  const attributes = value.attributes;
  if (!attributes || typeof attributes !== 'object' || Array.isArray(attributes)) return false;
  if (value.type === GALLERY_IMAGE_RECORD) {
    return validId(value.galleryId) && value.galleryId !== value.id && galleryImage(attributes);
  }
  if (value.type !== 'image_gallery') return false;
  const { layout, align, images, ...rest } = attributes as JsonObject;
  return (
    Object.keys(rest).length === 0 &&
    (layout === 'row-fit' || layout === 'grid') &&
    (align === undefined || align === 'left' || align === 'center' || align === 'right') &&
    // A projected cache may travel through ordinary payload-copy APIs. Native child order wins.
    (images === undefined ||
      (Array.isArray(images) && images.length <= 100 && images.every(galleryImage)))
  );
}

/** The array is a read projection. Field writes never store it as competing order state. */
export function galleryStoredAttributes(payload: JsonObject): Readonly<Record<string, JsonValue>> {
  const attributes = payload.attributes as JsonObject;
  if (payload.version !== GALLERY_RECORD_VERSION || payload.type !== 'image_gallery')
    return attributes;
  const { images: _images, ...stored } = attributes;
  return stored;
}
