import { REVIEW_GROUP_STORE_LIMITS } from './group-record-limits';
export { REVIEW_GROUP_STORE_LIMITS } from './group-record-limits';
import {
  customXmlDataParts,
  customXmlNodes,
  withCustomXmlDataPart,
  withCustomXmlNode,
  type CustomXmlDataPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  decodeReviewGroupRecord,
  encodeReviewGroupRecord,
  isReviewGroupRecordId,
  type ReviewGroupDecode,
  type ReviewGroupRecord,
} from './group-record-schema';
import {
  decodePayloadGroupDecisionRecord,
  encodePayloadGroupDecisionRecord,
  type PayloadGroupDecisionRecord,
} from './payload-group-decision-schema';

export const REVIEW_GROUP_STORE_NAMESPACE = 'urn:sciobot:docx:review-groups:1';
export const REVIEW_GROUP_STORE_ROOT = 'groups';

const REMOVED = 'removed';
export interface StoredReviewGroupRecord {
  readonly id: string;
  readonly raw: string;
  readonly decoded:
    | Exclude<ReviewGroupDecode, { status: 'invalid' }>
    | { readonly status: 'payload-decision'; readonly value: PayloadGroupDecisionRecord };
}
export type ReviewGroupStoreRead =
  | { readonly status: 'absent' }
  | { readonly status: 'invalid'; readonly reason: string }
  | {
      readonly status: 'ready';
      readonly part: CustomXmlDataPart;
      readonly records: ReadonlyMap<string, StoredReviewGroupRecord>;
      /** All authored records, including removed records retained for undo. */
      readonly retainedRecords: ReadonlyMap<string, StoredReviewGroupRecord>;
      readonly removedIds: ReadonlySet<string>;
    };
export type ReviewGroupStoreWrite =
  | { readonly ok: true; readonly pkg: OoxmlPackage }
  | { readonly ok: false; readonly pkg: OoxmlPackage; readonly reason: string };

/** One independent record per native node. The store is not a bound application block payload. */
export function readReviewGroupStore(pkg: OoxmlPackage): ReviewGroupStoreRead {
  const parts = customXmlDataParts(pkg, pkg.mainDocumentPart).filter(
    (part) => part.namespaceUri === REVIEW_GROUP_STORE_NAMESPACE
  );
  if (!parts.length) return { status: 'absent' };
  if (parts.length !== 1) return { status: 'invalid', reason: 'duplicate-review-group-store' };
  const descriptor = parts[0]!,
    part = pkg.parts.get(descriptor.partName);
  if (
    !part ||
    part.root.namespaceUri !== REVIEW_GROUP_STORE_NAMESPACE ||
    part.root.localName !== REVIEW_GROUP_STORE_ROOT
  )
    return { status: 'invalid', reason: 'invalid-review-group-store-root' };
  if (part.root.children.length > REVIEW_GROUP_STORE_LIMITS.maxNodeSlots)
    return { status: 'invalid', reason: 'review-group-store-node-limit' };
  const authored = part.root.children.filter(
      (node) =>
        node.kind !== 'textValue' &&
        node.namespaceUri === REVIEW_GROUP_STORE_NAMESPACE &&
        node.localName === 'node' &&
        (node.attributes.length || node.children.length)
    ),
    nodes = customXmlNodes(pkg, descriptor.partName);
  if (nodes.length !== authored.length)
    return { status: 'invalid', reason: 'missing-review-group-node-id' };
  if (nodes.length > REVIEW_GROUP_STORE_LIMITS.maxRecords)
    return { status: 'invalid', reason: 'review-group-store-count-limit' };
  for (const node of authored) {
    if (node.kind === 'textValue') continue;
    for (const name of ['label', 'data']) {
      const fields = node.children.filter(
        (child) =>
          child.kind !== 'textValue' &&
          child.namespaceUri === REVIEW_GROUP_STORE_NAMESPACE &&
          child.localName === name
      );
      if (
        fields.length !== 1 ||
        fields.some(
          (field) =>
            field.kind !== 'textValue' && field.children.some((child) => child.kind !== 'textValue')
        )
      )
        return { status: 'invalid', reason: 'invalid-review-group-node-fields' };
    }
  }
  let bytes = 0;
  const records = new Map<string, StoredReviewGroupRecord>(),
    retainedRecords = new Map<string, StoredReviewGroupRecord>(),
    removedIds = new Set<string>();
  for (const node of nodes) {
    if (!isReviewGroupRecordId(node.id) || records.has(node.id) || removedIds.has(node.id))
      return { status: 'invalid', reason: 'invalid-or-duplicate-review-group-node-id' };
    if (node.label !== '' && node.label !== REMOVED)
      return { status: 'invalid', reason: 'invalid-review-group-node-label' };
    bytes += new TextEncoder().encode(node.data).byteLength;
    if (bytes > REVIEW_GROUP_STORE_LIMITS.maxBytes)
      return { status: 'invalid', reason: 'review-group-store-byte-limit' };
    const native = decodeReviewGroupRecord(node.data);
    let decoded: StoredReviewGroupRecord['decoded'];
    if (native.status === 'unsupported-version' && native.version === 2) {
      const decision = decodePayloadGroupDecisionRecord(node.data);
      if (decision.status === 'invalid')
        return { status: 'invalid', reason: `${node.id}:${decision.reason}` };
      if (node.label !== '')
        return { status: 'invalid', reason: 'invalid-payload-group-decision-label' };
      decoded = { status: 'payload-decision', value: decision.value };
    } else {
      if (native.status === 'invalid')
        return { status: 'invalid', reason: `${node.id}:${native.reason}` };
      decoded = native;
    }
    if ((decoded.status === 'unsupported-version' ? decoded.id : decoded.value.id) !== node.id)
      return { status: 'invalid', reason: 'mismatched-review-group-node-id' };
    const entry = Object.freeze({ id: node.id, raw: node.data, decoded });
    retainedRecords.set(node.id, entry);
    if (node.label === REMOVED && decoded.status === 'supported') removedIds.add(node.id);
    else records.set(node.id, entry);
  }
  return { status: 'ready', part: descriptor, records, retainedRecords, removedIds };
}

/** Import/room initialization only. Live writes require this part to exist already. */
export function withInitializedReviewGroupStore(pkg: OoxmlPackage): ReviewGroupStoreWrite {
  const current = readReviewGroupStore(pkg);
  if (current.status === 'invalid') return { ok: false, pkg, reason: current.reason };
  if (current.status === 'ready') return { ok: true, pkg };
  const created = withCustomXmlDataPart(
    pkg,
    pkg.mainDocumentPart,
    REVIEW_GROUP_STORE_NAMESPACE,
    REVIEW_GROUP_STORE_ROOT
  );
  if (!created.part || readReviewGroupStore(created.pkg).status !== 'ready')
    return { ok: false, pkg, reason: 'review-group-store-initialization-refused' };
  return { ok: true, pkg: created.pkg };
}

/** Immutable records avoid concurrent replacement of a shared JSON string. */
export function withReviewGroupRecord(
  pkg: OoxmlPackage,
  record: ReviewGroupRecord
): ReviewGroupStoreWrite {
  const fail = (reason: string): ReviewGroupStoreWrite => ({ ok: false, pkg, reason });
  let raw: string;
  try {
    raw = encodeReviewGroupRecord(record);
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause));
  }
  return withStoredReviewGroupRecord(pkg, record.id, raw);
}

/** Fresh IDs let concurrent decisions retain independent history, including same-group conflicts. */
export function withPayloadGroupDecisionRecord(
  pkg: OoxmlPackage,
  record: PayloadGroupDecisionRecord
): ReviewGroupStoreWrite {
  let raw: string;
  try {
    raw = encodePayloadGroupDecisionRecord(record);
  } catch (cause) {
    return { ok: false, pkg, reason: cause instanceof Error ? cause.message : String(cause) };
  }
  return withStoredReviewGroupRecord(pkg, record.id, raw);
}

function withStoredReviewGroupRecord(
  pkg: OoxmlPackage,
  id: string,
  raw: string
): ReviewGroupStoreWrite {
  const fail = (reason: string): ReviewGroupStoreWrite => ({ ok: false, pkg, reason });
  const store = readReviewGroupStore(pkg);
  if (store.status !== 'ready')
    return fail(
      store.status === 'invalid' ? store.reason : 'review-group-store-initialization-required'
    );
  if (store.removedIds.has(id)) return fail('review-group-id-reserved');
  const previous = store.records.get(id);
  if (previous?.raw === raw) return { ok: true, pkg };
  if (previous)
    return fail(
      previous.decoded.status === 'unsupported-version'
        ? 'unsupported-review-group-version'
        : 'review-group-record-is-immutable'
    );
  if (store.records.size + store.removedIds.size >= REVIEW_GROUP_STORE_LIMITS.maxRecords)
    return fail('review-group-store-count-limit');
  const bytes = [...customXmlNodes(pkg, store.part.partName).map((node) => node.data), raw].reduce(
    (sum, data) => sum + new TextEncoder().encode(data).byteLength,
    0
  );
  if (bytes > REVIEW_GROUP_STORE_LIMITS.maxBytes) return fail('review-group-store-byte-limit');
  const next = withCustomXmlNode(pkg, store.part.partName, { id, label: '', data: raw }),
    read = readReviewGroupStore(next);
  if (read.status !== 'ready' || read.records.get(id)?.raw !== raw)
    return fail('review-group-record-write-refused');
  return { ok: true, pkg: next };
}

/** Reserve the stable ID and preserve field identity for peer-preserving removal undo. */
export function withoutReviewGroupRecord(pkg: OoxmlPackage, id: string): ReviewGroupStoreWrite {
  const fail = (reason: string): ReviewGroupStoreWrite => ({ ok: false, pkg, reason });
  if (!isReviewGroupRecordId(id)) return fail('invalid-review-group-id');
  const store = readReviewGroupStore(pkg);
  if (store.status === 'invalid') return fail(store.reason);
  if (store.status === 'absent' || !store.records.has(id)) return { ok: true, pkg };
  const record = store.records.get(id)!;
  if (record.decoded.status === 'payload-decision') return fail('review-group-record-is-immutable');
  if (record.decoded.status !== 'supported') return fail('unsupported-review-group-version');
  const next = withCustomXmlNode(pkg, store.part.partName, {
      id,
      label: REMOVED,
      data: record.raw,
    }),
    read = readReviewGroupStore(next);
  if (read.status !== 'ready' || !read.removedIds.has(id))
    return fail('review-group-record-removal-refused');
  return { ok: true, pkg: next };
}
