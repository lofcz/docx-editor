import {
  customXmlDataParts,
  customXmlNodes,
  replaceNode,
  withCustomXmlDataPart,
  withCustomXmlNode,
  withPart,
  type CustomXmlDataPart,
  type OoxmlElement,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { fieldChanges, readFields, withFieldChanges } from './payload-fields';
import { isPayloadLabel, readPayloadLabel, withPayloadLabel } from './payload-label';
import {
  BLOCK_TAG_PREFIX,
  PROPERTY_REVIEW_NODE_PREFIX,
  decodeBlockPayload,
  encodeBlockPayload,
  isBlockId,
  type BlockPayload,
  type PayloadDecodeResult,
} from './payload-schema';

export const PAYLOAD_NAMESPACE = 'urn:sciobot:docx:blocks:1';
export const PAYLOAD_ROOT = 'blocks';
export const PAYLOAD_NODE_PREFIX = 'scio_';
const MAX_STORE_NODES = 10000;
/** Word limits a content-control tag to 64 characters. */
const MAX_TAG_LENGTH = 64;

/** Where one embedder keeps its block payloads: names from its component registry storage. */
export interface PayloadStorage {
  readonly tagPrefix: string;
  readonly namespace: string;
  readonly root: string;
  readonly nodePrefix: string;
  readonly reviewNodePrefix: string;
}

/** The Sciobot storage names. Documents saved with them keep opening unchanged. */
export const SCIOBOT_PAYLOAD_STORAGE: PayloadStorage = Object.freeze({
  tagPrefix: BLOCK_TAG_PREFIX,
  namespace: PAYLOAD_NAMESPACE,
  root: PAYLOAD_ROOT,
  nodePrefix: PAYLOAD_NODE_PREFIX,
  reviewNodePrefix: PROPERTY_REVIEW_NODE_PREFIX,
});

export function blockTag(id: string, storage: PayloadStorage = SCIOBOT_PAYLOAD_STORAGE): string {
  if (!isBlockId(id) || storage.tagPrefix.length + id.length > MAX_TAG_LENGTH)
    throw new TypeError('Invalid block ID');
  return storage.tagPrefix + id;
}

export function blockIdFromTag(
  tag: string,
  storage: PayloadStorage = SCIOBOT_PAYLOAD_STORAGE
): string | null {
  if (!tag.startsWith(storage.tagPrefix)) return null;
  const id = tag.slice(storage.tagPrefix.length);
  return isBlockId(id) ? id : null;
}

export function payloadReferences(
  tags: Iterable<string>,
  storage: PayloadStorage = SCIOBOT_PAYLOAD_STORAGE
):
  | { readonly ok: true; readonly ids: ReadonlySet<string> }
  | { readonly ok: false; readonly reason: string } {
  const ids = new Set<string>();
  for (const tag of tags) {
    const id = blockIdFromTag(tag, storage);
    if (id === null) {
      if (tag.startsWith(storage.tagPrefix)) return { ok: false, reason: 'malformed-reference' };
      continue;
    }
    if (ids.has(id)) return { ok: false, reason: `duplicate-reference:${id}` };
    ids.add(id);
  }
  return { ok: true, ids };
}

/** Store `node` elements by their authored `id` attribute. */
function recordElements(part: OoxmlPart): Map<string, OoxmlElement> {
  const found = new Map<string, OoxmlElement>();
  for (const child of part.root.children) {
    if (child.kind === 'textValue' || child.localName !== 'node') continue;
    if (child.namespaceUri !== part.root.namespaceUri) continue;
    const id = child.attributes.find(
      (attribute) => attribute.localName === 'id' && attribute.namespaceUri === ''
    )?.value;
    if (id && !found.has(id)) found.set(id, child);
  }
  return found;
}

/** The record's effective JSON: its created `data` with field registers applied. */
function effectiveRaw(
  element: OoxmlElement | undefined,
  data: string
): { readonly ok: true; readonly raw: string } | { readonly ok: false; readonly reason: string } {
  if (!element) return { ok: true, raw: data };
  const fields = readFields(element);
  if (!fields.ok) return fields;
  if (!fields.count) return { ok: true, raw: data };
  let base: unknown;
  try {
    base = JSON.parse(data);
  } catch {
    return { ok: false, reason: 'invalid-json' };
  }
  return { ok: true, raw: JSON.stringify(fields.apply(base)) };
}

export interface StoredPayload {
  readonly id: string;
  readonly label: string;
  readonly raw: string;
  readonly decoded: Exclude<PayloadDecodeResult, { status: 'invalid' }>;
}

export type PayloadStoreRead =
  | { readonly status: 'absent' }
  | { readonly status: 'invalid'; readonly reason: string }
  | {
      readonly status: 'ready';
      readonly part: CustomXmlDataPart;
      readonly entries: ReadonlyMap<string, StoredPayload>;
    };

/** Initialize before room bootstrap so no actor owns the shared store lifetime. */
export function withInitializedPayloadStore(
  pkg: OoxmlPackage,
  storage: PayloadStorage = SCIOBOT_PAYLOAD_STORAGE
):
  | { readonly ok: true; readonly pkg: OoxmlPackage }
  | { readonly ok: false; readonly pkg: OoxmlPackage; readonly reason: string } {
  const current = readPayloadStore(pkg, storage);
  if (current.status === 'invalid') return { ok: false, pkg, reason: current.reason };
  if (current.status === 'ready') return { ok: true, pkg };
  const created = withCustomXmlDataPart(pkg, pkg.mainDocumentPart, storage.namespace, storage.root);
  if (!created.part || readPayloadStore(created.pkg, storage).status !== 'ready')
    return { ok: false, pkg, reason: 'payload-store-initialization-refused' };
  return { ok: true, pkg: created.pkg };
}

export function readPayloadStore(
  pkg: OoxmlPackage,
  storage: PayloadStorage = SCIOBOT_PAYLOAD_STORAGE
): PayloadStoreRead {
  const stores = customXmlDataParts(pkg, pkg.mainDocumentPart).filter(
    (part) => part.namespaceUri === storage.namespace
  );
  if (!stores.length) return { status: 'absent' };
  if (stores.length !== 1) return { status: 'invalid', reason: 'duplicate-store' };
  const descriptor = stores[0]!;
  const part = pkg.parts.get(descriptor.partName);
  if (!part || part.root.localName !== storage.root)
    return { status: 'invalid', reason: 'invalid-store-root' };
  if (part.root.children.length > 50_000) return { status: 'invalid', reason: 'store-node-limit' };
  const nodes = customXmlNodes(pkg, descriptor.partName);
  const authoredCount = part.root.children.filter(
    (node) =>
      node.kind !== 'textValue' &&
      node.namespaceUri === storage.namespace &&
      node.localName === 'node' &&
      // Peer-preserving undo can retain an empty native wrapper without a payload or identity.
      (node.attributes.length !== 0 || node.children.length !== 0)
  ).length;
  if (nodes.length !== authoredCount) return { status: 'invalid', reason: 'missing-node-id' };
  if (nodes.length > MAX_STORE_NODES) return { status: 'invalid', reason: 'store-node-limit' };
  const entries = new Map<string, StoredPayload>();
  const elements = recordElements(part);
  for (const node of nodes) {
    // This separate record family has its own bounded review reader and no block binding.
    if (node.id.startsWith(storage.reviewNodePrefix)) continue;
    const id = node.id.startsWith(storage.nodePrefix)
      ? node.id.slice(storage.nodePrefix.length)
      : null;
    if (!isBlockId(id)) return { status: 'invalid', reason: 'invalid-node-id' };
    if (entries.has(id)) return { status: 'invalid', reason: `duplicate-node:${id}` };
    const effective = effectiveRaw(elements.get(node.id), node.data);
    if (!effective.ok) return { status: 'invalid', reason: `${id}:${effective.reason}` };
    const decoded = decodeBlockPayload(effective.raw);
    if (decoded.status === 'invalid')
      return { status: 'invalid', reason: `${id}:${decoded.reason}` };
    if (decoded.status === 'supported' && decoded.value.id !== id)
      return { status: 'invalid', reason: `mismatched-node-id:${id}` };
    // Opaque future payloads retain their own label semantics as well as raw metadata.
    const label =
      decoded.status === 'supported'
        ? readPayloadLabel(elements.get(node.id), node.label)
        : { ok: true as const, label: node.label };
    if (!label.ok) return { status: 'invalid', reason: `${id}:${label.reason}` };
    entries.set(id, { id, label: label.label, raw: effective.raw, decoded });
  }
  return { status: 'ready', part: descriptor, entries };
}

export type PayloadStoreWrite =
  | { readonly ok: false; readonly pkg: OoxmlPackage; readonly reason: string }
  | {
      readonly ok: true;
      readonly pkg: OoxmlPackage;
      readonly part: CustomXmlDataPart;
      readonly tag: string;
    };

/** Build an immutable package candidate. The live editor transaction belongs to B02. */
export function withPayloadRecord(
  pkg: OoxmlPackage,
  payload: BlockPayload,
  label: string,
  storage: PayloadStorage = SCIOBOT_PAYLOAD_STORAGE
): PayloadStoreWrite {
  const refuse = (reason: string): PayloadStoreWrite => ({ ok: false, pkg, reason });
  if (!isPayloadLabel(label)) return refuse('invalid-label');
  const existing = readPayloadStore(pkg, storage);
  if (existing.status === 'invalid') return refuse(existing.reason);
  if (
    existing.status === 'ready' &&
    existing.entries.get(payload.id)?.decoded.status === 'unsupported-version'
  )
    return refuse('unsupported-version');
  let raw: string;
  try {
    raw = encodeBlockPayload(payload);
  } catch (error) {
    return refuse(error instanceof Error ? error.message : 'invalid-payload');
  }
  const current = existing.status === 'ready' ? existing.entries.get(payload.id) : undefined;
  if (existing.status === 'ready' && current?.decoded.status === 'supported') {
    const updated = withFieldRegisters(
      pkg,
      existing,
      current.decoded.value,
      current.label,
      payload,
      label,
      storage
    );
    if (!updated) return refuse('payload-write-refused');
    const reopened = readPayloadStore(updated, storage);
    const stored = reopened.status === 'ready' ? reopened.entries.get(payload.id) : undefined;
    if (
      stored?.decoded.status !== 'supported' ||
      stored.label !== label ||
      fieldChanges(stored.decoded.value, payload).length
    )
      return refuse('payload-write-refused');
    return { ok: true, pkg: updated, part: existing.part, tag: blockTag(payload.id, storage) };
  }
  const created = withCustomXmlDataPart(pkg, pkg.mainDocumentPart, storage.namespace, storage.root);
  if (!created.part) return refuse('store-creation-refused');
  const next = withCustomXmlNode(created.pkg, created.part.partName, {
    id: storage.nodePrefix + payload.id,
    label,
    data: raw,
  });
  const reopened = readPayloadStore(next, storage);
  if (reopened.status !== 'ready' || reopened.entries.get(payload.id)?.raw !== raw)
    return refuse('payload-write-refused');
  return { ok: true, pkg: next, part: created.part, tag: blockTag(payload.id, storage) };
}

/**
 * Update an existing record through field registers, leaving its created `data` text intact.
 * A later effective label uses independent immutable scalar registers; created Word label text stays intact.
 */
function withFieldRegisters(
  pkg: OoxmlPackage,
  store: Extract<PayloadStoreRead, { status: 'ready' }>,
  current: BlockPayload,
  currentLabel: string,
  payload: BlockPayload,
  label: string,
  storage: PayloadStorage
): OoxmlPackage | null {
  const partName = store.part.partName;
  const nodeId = storage.nodePrefix + payload.id;
  const updates = fieldChanges(current, payload);
  if (!updates.length && currentLabel === label) return pkg;
  const part = pkg.parts.get(partName);
  const record = part && recordElements(part).get(nodeId);
  if (!part || !record) return null;
  const labeled = currentLabel === label ? record : withPayloadLabel(record, label);
  const used = new Set<string>();
  const collect = (node: OoxmlPart['root'] | OoxmlElement['children'][number]) => {
    used.add(node.id);
    if (node.kind !== 'textValue') node.children.forEach(collect);
  };
  collect(part.root);
  const replaced = replaceNode(part, record.id, {
    ...labeled,
    children: updates.length ? withFieldChanges(labeled, updates, used) : labeled.children,
  } as OoxmlElement);
  if (!replaced.ok) return null;
  return withPart(pkg, replaced.part);
}
