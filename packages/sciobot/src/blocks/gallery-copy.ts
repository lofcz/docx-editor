import type { BlockPayload } from './payload-schema';
import { GALLERY_IMAGE_RECORD } from './gallery-record-shape';

/** Remap only declared gallery ownership. Captions and other authored strings stay literal. */
export function copiedGalleryPayload(
  payload: BlockPayload,
  ids: ReadonlyMap<string, string>
): BlockPayload {
  if (payload.version !== 2 && payload.version !== 3) return payload;
  if (payload.type === 'image_gallery') {
    const { images: _images, ...attributes } = payload.attributes;
    return { ...payload, attributes };
  }
  if (payload.type === GALLERY_IMAGE_RECORD) {
    const galleryId = typeof payload.galleryId === 'string' && ids.get(payload.galleryId);
    if (!galleryId) throw Error('A copied gallery image requires its copied gallery owner.');
    return {
      ...payload,
      galleryId,
      ...(payload.version === 3 ? { attributes: { resource: payload.attributes.resource! } } : {}),
    };
  }
  return payload;
}

/** Child records exist before their parent exposes the ordered image projection. */
export function galleryRecordOrder(a: BlockPayload, b: BlockPayload): number {
  return (
    Number(a.version >= 2 && a.type === 'image_gallery') -
    Number(b.version >= 2 && b.type === 'image_gallery')
  );
}
