import {
  customXmlDataParts,
  customXmlNodes,
  withCustomXmlNode,
  type CustomXmlDataPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { isBlockId, PROPERTY_REVIEW_NODE_PREFIX } from '../blocks/payload-schema';
import { PAYLOAD_NAMESPACE, PAYLOAD_ROOT } from '../blocks/payload-opc-codec';
import {
  decodePropertyReviewRecord,
  encodePropertyReviewRecord,
  type PropertyRecordDecode,
  type PropertyReviewRecord,
} from './property-record-schema';

export const PROPERTY_REVIEW_NAMESPACE = PAYLOAD_NAMESPACE;
export const PROPERTY_REVIEW_ROOT = PAYLOAD_ROOT;
const PREFIX = PROPERTY_REVIEW_NODE_PREFIX;
const MAX_RECORDS = 1024;
const MAX_NODE_SLOTS = 50_000;
const MAX_STORE_BYTES = 8 * 1024 * 1024;
const REMOVED = 'removed';
export interface StoredPropertyRecord {
  readonly id: string;
  readonly raw: string;
  readonly decoded: Exclude<PropertyRecordDecode, { status: 'invalid' }>;
}
export type PropertyStoreRead =
  | { readonly status: 'absent' }
  | { readonly status: 'invalid'; readonly reason: string }
  | {
      readonly status: 'ready';
      readonly part: CustomXmlDataPart;
      readonly records: ReadonlyMap<string, StoredPropertyRecord>;
      /** Immutable receipts include removed proposals retained for either actor's decision undo. */
      readonly retainedRecords: ReadonlyMap<string, StoredPropertyRecord>;
      /** Resolved IDs remain reserved. Their native fields retain peer-preserving undo identity. */
      readonly removedIds: ReadonlySet<string>;
    };
export type PropertyStoreWrite =
  | { readonly ok: true; readonly pkg: OoxmlPackage }
  | { readonly ok: false; readonly pkg: OoxmlPackage; readonly reason: string };

/** One proposal per native custom XML node. Independent proposals never overwrite one payload string. */
export function readPropertyReviewStore(pkg: OoxmlPackage): PropertyStoreRead {
  const descriptors = customXmlDataParts(pkg, pkg.mainDocumentPart).filter(
    (part) => part.namespaceUri === PROPERTY_REVIEW_NAMESPACE
  );
  if (!descriptors.length) return { status: 'absent' };
  if (descriptors.length !== 1) return { status: 'invalid', reason: 'duplicate-property-store' };
  const descriptor = descriptors[0]!,
    part = pkg.parts.get(descriptor.partName);
  if (
    !part ||
    part.root.localName !== PROPERTY_REVIEW_ROOT ||
    part.root.namespaceUri !== PROPERTY_REVIEW_NAMESPACE
  )
    return { status: 'invalid', reason: 'invalid-property-store-root' };
  if (part.root.children.length > MAX_NODE_SLOTS)
    return { status: 'invalid', reason: 'property-store-node-limit' };
  const nodes = customXmlNodes(pkg, descriptor.partName),
    authored = part.root.children.filter(
      (node) =>
        node.kind !== 'textValue' &&
        node.namespaceUri === PROPERTY_REVIEW_NAMESPACE &&
        node.localName === 'node' &&
        // Native peer-preserving undo can leave an empty wrapper. It carries no record or intent.
        (node.attributes.length !== 0 || node.children.length !== 0)
    );
  if (nodes.length !== authored.length)
    return { status: 'invalid', reason: 'missing-property-node-id' };
  for (const node of authored) {
    if (node.kind === 'textValue') continue;
    const id = node.attributes.find(
      (attribute) => attribute.namespaceUri === '' && attribute.localName === 'id'
    )?.value;
    if (!id?.startsWith(PREFIX)) continue;
    for (const name of ['label', 'data']) {
      const fields = node.children.filter(
        (child) =>
          child.kind !== 'textValue' &&
          child.namespaceUri === PROPERTY_REVIEW_NAMESPACE &&
          child.localName === name
      );
      if (
        fields.length !== 1 ||
        fields.some(
          (field) =>
            field.kind !== 'textValue' && field.children.some((child) => child.kind !== 'textValue')
        )
      )
        return { status: 'invalid', reason: 'invalid-property-node-fields' };
    }
  }
  if (nodes.filter((node) => node.id.startsWith(PREFIX)).length > MAX_RECORDS)
    return { status: 'invalid', reason: 'property-store-count-limit' };
  const records = new Map<string, StoredPropertyRecord>(),
    retainedRecords = new Map<string, StoredPropertyRecord>(),
    removedIds = new Set<string>();
  let bytes = 0;
  for (const node of nodes) {
    if (node.id.startsWith('scio_')) continue;
    const id = node.id.startsWith(PREFIX) ? node.id.slice(PREFIX.length) : null;
    if (!isBlockId(id)) return { status: 'invalid', reason: 'invalid-property-node-id' };
    if (records.has(id) || removedIds.has(id))
      return { status: 'invalid', reason: 'duplicate-property-node-id' };
    if (node.label !== '' && node.label !== REMOVED)
      return { status: 'invalid', reason: 'unexpected-property-node-label' };
    bytes += new TextEncoder().encode(node.data).byteLength;
    if (bytes > MAX_STORE_BYTES) return { status: 'invalid', reason: 'property-store-byte-limit' };
    const decoded = decodePropertyReviewRecord(node.data);
    if (decoded.status === 'invalid')
      return { status: 'invalid', reason: `${id}:${decoded.reason}` };
    const dataId = decoded.status === 'supported' ? decoded.value.id : decoded.id;
    if (dataId !== id) return { status: 'invalid', reason: 'mismatched-property-node-id' };
    const entry = Object.freeze({ id, raw: node.data, decoded });
    retainedRecords.set(id, entry);
    if (node.label === REMOVED && decoded.status === 'supported') {
      removedIds.add(id);
      continue;
    }
    records.set(id, entry);
  }
  return { status: 'ready', part: descriptor, records, retainedRecords, removedIds };
}

/** Immutable candidate only. The caller must use one guarded native transaction for live changes. */
export function withPropertyReviewRecord(
  pkg: OoxmlPackage,
  record: PropertyReviewRecord
): PropertyStoreWrite {
  const refuse = (reason: string): PropertyStoreWrite => ({ ok: false, pkg, reason });
  let raw: string;
  try {
    raw = encodePropertyReviewRecord(record);
  } catch (error) {
    return refuse(error instanceof Error ? error.message : 'invalid-property-record');
  }
  const old = readPropertyReviewStore(pkg);
  if (old.status === 'invalid') return refuse(old.reason);
  if (old.status === 'ready' && old.removedIds.has(record.id))
    return refuse('removed-property-record-id');
  const entries = old.status === 'ready' ? old.records : new Map<string, StoredPropertyRecord>(),
    previous = entries.get(record.id);
  if (previous?.decoded.status === 'unsupported-version')
    return refuse('unsupported-property-record-version');
  if (previous?.raw === raw) return { ok: true, pkg };
  if (!previous && entries.size >= MAX_RECORDS) return refuse('property-store-count-limit');
  let bytes = new TextEncoder().encode(raw).byteLength;
  for (const [id, entry] of entries)
    if (id !== record.id) bytes += new TextEncoder().encode(entry.raw).byteLength;
  if (bytes > MAX_STORE_BYTES) return refuse('property-store-byte-limit');
  // Every property owner already has an application payload. Avoid racing a second part's creation.
  if (old.status !== 'ready') return refuse('application-payload-store-required');
  const next = withCustomXmlNode(pkg, old.part.partName, {
      id: PREFIX + record.id,
      label: '',
      data: raw,
    }),
    read = readPropertyReviewStore(next);
  if (read.status !== 'ready' || read.records.get(record.id)?.raw !== raw)
    return refuse('property-record-write-refused');
  return { ok: true, pkg: next };
}

/** Mark one stable node removed. Root child positions can differ after concurrent native materialization. */
export function withoutPropertyReviewRecord(pkg: OoxmlPackage, id: string): PropertyStoreWrite {
  const refuse = (reason: string): PropertyStoreWrite => ({ ok: false, pkg, reason });
  if (!isBlockId(id)) return refuse('invalid-property-record-id');
  const store = readPropertyReviewStore(pkg);
  if (store.status === 'invalid') return refuse(store.reason);
  if (store.status === 'absent' || !store.records.has(id)) return { ok: true, pkg };
  if (store.records.get(id)!.decoded.status === 'unsupported-version')
    return refuse('unsupported-property-record-version');
  const next = withCustomXmlNode(pkg, store.part.partName, {
      id: PREFIX + id,
      label: REMOVED,
      data: store.records.get(id)!.raw,
    }),
    read = readPropertyReviewStore(next);
  if (read.status !== 'ready' || read.records.has(id))
    return refuse('property-record-removal-refused');
  return { ok: true, pkg: next };
}
