import type { BlockPayload } from '../blocks/payload-schema';
import type { JsonObject, JsonValue } from '../data/json';

/** Content-addressed attachment paths. These are data references, never fetchable URLs. */
export const GALLERY_ASSET_PATH = /^media\/([a-f0-9]{64})\.(png|jpg|gif|bmp|webp)$/;
export interface GalleryAssetImage extends JsonObject {
  readonly asset: string;
  readonly width: number;
  readonly height: number;
  readonly caption: string;
  readonly alt: string;
  readonly title: string | null;
}

function assetImage(value: JsonValue): value is GalleryAssetImage {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const { asset, width, height, caption, alt, title, ...extra } = value as JsonObject;
  return (
    Object.keys(extra).length === 0 &&
    typeof asset === 'string' &&
    GALLERY_ASSET_PATH.test(asset) &&
    [width, height].every(
      (size) => typeof size === 'number' && Number.isFinite(size) && size > 0 && size <= 211200
    ) &&
    [caption, alt].every((text) => typeof text === 'string' && text.length <= 2048) &&
    (title === null || (typeof title === 'string' && title.length <= 2048))
  );
}

/** Versioned import syntax stays separate from the legacy HTTP gallery grammar. */
export function validGalleryAssetPayload(payload: BlockPayload): boolean {
  const { resourceVersion, layout, align, images, ...extra } = payload.attributes;
  return (
    payload.type === 'image_gallery' &&
    payload.version === 1 &&
    resourceVersion === '1' &&
    Object.keys(extra).length === 0 &&
    (layout === 'row-fit' || layout === 'grid') &&
    (align === undefined || align === 'left' || align === 'center' || align === 'right') &&
    Array.isArray(images) &&
    images.length > 0 &&
    images.length <= 100 &&
    images.every(assetImage)
  );
}
