import {
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  contentControlsIn,
  createNodeIdAllocator,
  isStoryPart,
  paragraphTextOf,
  replaceChildren,
  replaceNode,
  withCustomXmlNode,
  withPart,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { JsonObject } from '../data/json';
import { validGallery } from '../projection/gallery';
import { activeMdxRegistry } from '../registry/active';
import { galleryImageControl, galleryImageText } from './gallery-record-build';
import { GALLERY_IMAGE_RECORD } from './gallery-record-shape';
import { blockIdFromTag, readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { encodeBlockPayload, type BlockPayload } from './payload-schema';

/**
 * Upgrade before room bootstrap, or only a newly created gallery in its insertion transaction.
 * Never call this on legacy records already owned by a live collaboration room.
 */
export function withGalleryImageRecords(
  pkg: OoxmlPackage,
  onlyIds?: ReadonlySet<string>
): { ok: true; pkg: OoxmlPackage } | { ok: false; pkg: OoxmlPackage; reason: string } {
  const store = readPayloadStore(pkg);
  const refuse = (reason: string) => ({ ok: false as const, pkg, reason });
  if (store.status === 'invalid') return refuse(store.reason);
  if (store.status !== 'ready') return { ok: true, pkg };
  const storage = activeMdxRegistry().storage;
  const ids = new Set(store.entries.keys());
  let next = pkg;
  for (const record of store.entries.values()) {
    if (onlyIds && !onlyIds.has(record.id)) continue;
    if (record.decoded.status !== 'supported') continue;
    const payload = record.decoded.value;
    if (payload.type !== 'image_gallery' || payload.version !== 1) continue;
    if (!validGallery(payload)) return refuse(`invalid-legacy-gallery:${record.id}`);
    const matches = [...next.parts.values()].filter(isStoryPart).flatMap((part) =>
      contentControlsIn(part.root)
        .filter(
          ({ node }) => blockIdFromTag(contentControlPropertiesOf(node).tag ?? '') === record.id
        )
        .map(({ node }) => ({ part, node }))
    );
    // Preserve an unbound historical payload. It has no image controls to upgrade.
    if (!matches.length) continue;
    if (matches.length !== 1) return refuse(`ambiguous-gallery-control:${record.id}`);
    const { part, node } = matches[0]!;
    const content = contentControlContentNodeOf(node);
    const images = payload.attributes.images as JsonObject[];
    if (
      !content ||
      content.children.length !== images.length ||
      content.children.some(
        (child, index) =>
          child.kind !== 'paragraph' ||
          paragraphTextOf(part, child.id) !== galleryImageText(images[index]!)
      )
    )
      return refuse(`modified-legacy-gallery-fallback:${record.id}`);
    const allocate = createNodeIdAllocator(part, 'new');
    let serial = 1;
    const children = images.map((image, index) => {
      let imageId: string;
      do {
        imageId = `${record.id.slice(0, 40)}-image-${serial++}`;
      } while (ids.has(imageId));
      ids.add(imageId);
      const item: BlockPayload = {
        id: imageId,
        type: GALLERY_IMAGE_RECORD,
        version: 2,
        galleryId: record.id,
        attributes: image,
      };
      const stored = withPayloadRecord(next, item, galleryImageText(image));
      if (!stored.ok) throw Error(stored.reason);
      next = stored.pkg;
      return galleryImageControl(part, imageId, image, content.children[index], allocate);
    });
    const changed = replaceChildren(part, content.id, children);
    if (!changed.ok) return refuse(`gallery-control-upgrade-refused:${record.id}`);
    next = withPart(next, changed.part);
    const { images: _images, ...attributes } = payload.attributes;
    const upgraded: BlockPayload = { ...payload, version: 2, attributes };
    // Bootstrap owns this rewrite. Existing field overrides must not reintroduce version-1 arrays.
    const dataPart = next.parts.get(store.part.partName)!;
    const dataNode = dataPart.root.children.find(
      (child) =>
        child.kind !== 'textValue' &&
        child.attributes.some(
          (attribute) =>
            attribute.namespaceUri === '' &&
            attribute.localName === 'id' &&
            attribute.value === storage.nodePrefix + record.id
        )
    );
    if (!dataNode || dataNode.kind === 'textValue') return refuse('missing-gallery-payload');
    const clean = replaceNode(dataPart, dataNode.id, {
      ...dataNode,
      children: dataNode.children.filter(
        (child) =>
          child.kind !== 'textValue' &&
          child.namespaceUri === dataNode.namespaceUri &&
          ['data', 'label'].includes(child.localName)
      ),
    } as OoxmlElement);
    if (!clean.ok) return refuse('gallery-payload-upgrade-refused');
    next = withCustomXmlNode(withPart(next, clean.part), store.part.partName, {
      id: storage.nodePrefix + record.id,
      label: record.label,
      data: encodeBlockPayload(upgraded),
    });
  }
  const verified = readPayloadStore(next);
  return verified.status === 'ready' ? { ok: true, pkg: next } : refuse('gallery-upgrade-refused');
}
