import {
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  resolveImageRelationship,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { decodeInlineImagePayload } from './inline-image-codec';
import { encodeBlockPayload } from './payload-schema';
import type { PayloadStorage, StoredPayload } from './payload-opc-codec';

const attribute = (node: OoxmlElement, name: string) =>
  node.attributes.find((entry) => entry.namespaceUri === '' && entry.localName === name)?.value;

/** Read the single native picture owned by a managed inline control. */
export function nativeImageFields(control: OoxmlNode) {
  const drawings: OoxmlElement[] = [];
  const metadata: OoxmlElement[] = [];
  const extents: OoxmlElement[] = [];
  let nested = false;
  const pending = [control];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'textValue') continue;
    if (node !== control && node.kind === 'contentControl') nested = true;
    if (node.kind === 'drawing') drawings.push(node);
    if (node.kind === 'drawingDocPr') metadata.push(node);
    if (node.kind === 'drawingExtent') extents.push(node);
    for (const child of node.children) pending.push(child);
  }
  if (nested || drawings.length !== 1 || metadata.length !== 1 || extents.length !== 1) return null;
  const width = Number(attribute(extents[0]!, 'cx'));
  const height = Number(attribute(extents[0]!, 'cy'));
  if (![width, height].every((value) => Number.isSafeInteger(value) && value > 0)) return null;
  return {
    drawing: drawings[0]!,
    attributes: {
      alt: attribute(metadata[0]!, 'descr') ?? '',
      title: attribute(metadata[0]!, 'title') || null,
      widthPoints: width / 12700,
      heightPoints: height / 12700,
    },
  };
}

/** Resolve one embedded native resource without loading external image URLs. */
export function nativeImageResource(pkg: OoxmlPackage, partName: string, control: OoxmlNode) {
  const native = nativeImageFields(control);
  if (!native) return null;
  const references: string[] = [];
  let linked = false;
  const pending: OoxmlNode[] = [native.drawing];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'textValue') continue;
    for (const attribute of node.attributes) {
      if (
        attribute.namespaceUri !==
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
      )
        continue;
      if (attribute.localName === 'embed') references.push(attribute.value);
      if (attribute.localName === 'link') linked = true;
    }
    for (const child of node.children) pending.push(child);
  }
  if (linked || references.length !== 1) return null;
  const resource = resolveImageRelationship(
    pkg.relationships.get(partName),
    partName,
    references[0]!
  );
  return resource.mode === 'internal' ? resource : null;
}

/** Native fields have one owner. Private records retain resource provenance and opaque metadata. */
export function projectNativeImageFields(
  pkg: OoxmlPackage,
  entries: Map<string, StoredPayload>,
  storage: PayloadStorage
): void {
  const images = new Set(
    [...entries.values()]
      .filter(
        (entry) =>
          entry.decoded.status === 'supported' && entry.decoded.value.type === 'inline_image'
      )
      .map((entry) => entry.id)
  );
  if (!images.size) return;
  const owners = new Map<string, { node: OoxmlNode; partName: string } | null>();
  for (const part of pkg.parts.values()) {
    if (!isStoryPart(part)) continue;
    for (const { node } of contentControlsIn(part.root)) {
      const tag = contentControlPropertiesOf(node).tag;
      if (!tag?.startsWith(storage.tagPrefix)) continue;
      const id = tag.slice(storage.tagPrefix.length);
      if (images.has(id)) owners.set(id, owners.has(id) ? null : { node, partName: part.name });
    }
  }
  for (const id of images) {
    const record = entries.get(id)!;
    if (record.decoded.status !== 'supported') continue;
    const payload = decodeInlineImagePayload(encodeBlockPayload(record.decoded.value));
    const owner = owners.get(id);
    const native = owner && nativeImageFields(owner.node);
    const resource = owner && nativeImageResource(pkg, owner.partName, owner.node);
    // Unbound, ambiguous and opaque records retain their authored metadata.
    if (!payload || !native || !resource || resource.partName !== payload.attributes.mediaPartName)
      continue;
    entries.set(id, {
      ...record,
      decoded: {
        status: 'supported',
        value: { ...payload, attributes: { ...payload.attributes, ...native.attributes } },
      },
    });
  }
}
