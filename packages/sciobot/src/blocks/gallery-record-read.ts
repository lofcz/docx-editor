import {
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  isAtomicContentControl,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { JsonObject } from '../data/json';
import type { StoredPayload, PayloadStorage } from './payload-opc-codec';
import { GALLERY_IMAGE_RECORD } from './gallery-record-shape';
import { readNativeGalleryImage } from './gallery-native-record';

/** The native children own image order. Metadata records own each image's fields. */
export function projectGalleryRecords(
  pkg: OoxmlPackage,
  entries: Map<string, StoredPayload>,
  storage: PayloadStorage
): string | null {
  const galleries = [...entries.values()].filter(
    (record) =>
      record.decoded.status === 'supported' &&
      [2, 3].includes(record.decoded.value.version) &&
      record.decoded.value.type === 'image_gallery'
  );
  if (!galleries.length) return null;
  const controls = new Map<string, { node: OoxmlNode; part: OoxmlPart }[]>();
  for (const part of pkg.parts.values()) {
    if (!isStoryPart(part)) continue;
    for (const { node } of contentControlsIn(part.root)) {
      const tag = contentControlPropertiesOf(node).tag;
      if (!tag?.startsWith(storage.tagPrefix)) continue;
      const id = tag.slice(storage.tagPrefix.length);
      const previous = controls.get(id);
      if (previous) previous.push({ node, part });
      else controls.set(id, [{ node, part }]);
    }
  }
  for (const record of galleries) {
    if (record.decoded.status !== 'supported') continue;
    const payload = record.decoded.value;
    const bound = controls.get(record.id) ?? [];
    // Unbound metadata can remain during native removal history. It owns no visible projection.
    if (!bound.length) continue;
    if (bound.length !== 1) return `ambiguous-gallery-control:${record.id}`;
    const content = contentControlContentNodeOf(bound[0]!.node);
    if (!content) return `missing-gallery-content:${record.id}`;
    const images: JsonObject[] = [];
    const seen = new Set<string>();
    for (const child of content.children) {
      if (payload.version === 2 ? !isAtomicContentControl(child) : child.kind !== 'contentControl')
        return `invalid-gallery-image-owner:${record.id}`;
      const tag = contentControlPropertiesOf(child).tag;
      if (!tag?.startsWith(storage.tagPrefix)) return `invalid-gallery-image-tag:${record.id}`;
      const id = tag.slice(storage.tagPrefix.length);
      const image = entries.get(id)?.decoded;
      if (
        !image ||
        image.status !== 'supported' ||
        image.value.type !== GALLERY_IMAGE_RECORD ||
        image.value.version !== payload.version ||
        image.value.galleryId !== record.id ||
        seen.has(id) ||
        controls.get(id)?.length !== 1
      )
        return `invalid-gallery-image-reference:${record.id}:${id}`;
      seen.add(id);
      if (images.length >= 100) return `gallery-image-limit:${record.id}`;
      let attributes = image.value.attributes;
      if (payload.version === 3) {
        const native = readNativeGalleryImage(pkg, bound[0]!.part, child);
        if (!native) return `invalid-native-gallery-image:${record.id}:${id}`;
        attributes = native.attributes;
        entries.set(id, {
          ...entries.get(id)!,
          decoded: { status: 'supported', value: { ...image.value, attributes } },
        });
      }
      images.push({ ...attributes });
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
