import {
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  isAtomicContentControl,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { JsonObject } from '../data/json';
import type { StoredPayload, PayloadStorage } from './payload-opc-codec';
import { GALLERY_IMAGE_RECORD } from './gallery-record-shape';

/** The native children own image order. Metadata records own each image's fields. */
export function projectGalleryRecords(
  pkg: OoxmlPackage,
  entries: Map<string, StoredPayload>,
  storage: PayloadStorage
): string | null {
  const galleries = [...entries.values()].filter(
    (record) =>
      record.decoded.status === 'supported' &&
      record.decoded.value.version === 2 &&
      record.decoded.value.type === 'image_gallery'
  );
  if (!galleries.length) return null;
  const controls = new Map<string, OoxmlNode[]>();
  for (const part of pkg.parts.values()) {
    if (!isStoryPart(part)) continue;
    for (const { node } of contentControlsIn(part.root)) {
      const tag = contentControlPropertiesOf(node).tag;
      if (!tag?.startsWith(storage.tagPrefix)) continue;
      const id = tag.slice(storage.tagPrefix.length);
      const previous = controls.get(id);
      if (previous) previous.push(node);
      else controls.set(id, [node]);
    }
  }
  for (const record of galleries) {
    if (record.decoded.status !== 'supported') continue;
    const payload = record.decoded.value;
    const bound = controls.get(record.id) ?? [];
    // Unbound metadata can remain during native removal history. It owns no visible projection.
    if (!bound.length) continue;
    if (bound.length !== 1) return `ambiguous-gallery-control:${record.id}`;
    const content = contentControlContentNodeOf(bound[0]!);
    if (!content) return `missing-gallery-content:${record.id}`;
    const images: JsonObject[] = [];
    const seen = new Set<string>();
    for (const child of content.children) {
      if (!isAtomicContentControl(child)) return `invalid-gallery-image-owner:${record.id}`;
      const tag = contentControlPropertiesOf(child).tag;
      if (!tag?.startsWith(storage.tagPrefix)) return `invalid-gallery-image-tag:${record.id}`;
      const id = tag.slice(storage.tagPrefix.length);
      const image = entries.get(id)?.decoded;
      if (
        !image ||
        image.status !== 'supported' ||
        image.value.type !== GALLERY_IMAGE_RECORD ||
        image.value.version !== 2 ||
        image.value.galleryId !== record.id ||
        seen.has(id) ||
        controls.get(id)?.length !== 1
      )
        return `invalid-gallery-image-reference:${record.id}:${id}`;
      seen.add(id);
      if (images.length >= 100) return `gallery-image-limit:${record.id}`;
      images.push({ ...image.value.attributes });
    }
    if (!images.length) return `empty-gallery-record:${record.id}`;
    entries.set(record.id, {
      ...record,
      decoded: {
        status: 'supported',
        value: {
          ...payload,
          attributes: { ...payload.attributes, images },
        },
      },
    });
  }
  return null;
}
