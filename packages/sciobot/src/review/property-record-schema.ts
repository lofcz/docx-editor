export const FIELD_PROPERTY_REVIEW_EXECUTOR = 'sciobot-field-properties-v1';
import type { RevisionAddress } from '@docx-editor.dev/core/store';
import { DEFAULT_PAYLOAD_LIMITS, isBlockId } from '../blocks/payload-schema';
import { copyBoundedJson, type JsonObject, type JsonValue } from '../data/json';
import type { PlannedPropertyChange, PlannedValue } from '../diff/plan';
import { jsonDataKey } from '../diff/target-key';

/** Root replacement slot, interpreted only by the sciobot-atomic-v1 executor. */
export const ATOMIC_PAYLOAD_CHANGE = '$payload';
export const ATOMIC_REVIEW_EXECUTOR = 'sciobot-atomic-v1';

export interface StoredPropertyBlock {
  readonly blockId: string;
  readonly type: string;
  /** Attribute names for property proposals; one $payload pair for sciobot-atomic-v1. */
  readonly changes: readonly PlannedPropertyChange[];
}
export interface StoredPropertyRevision extends RevisionAddress {
  readonly partName: string;
}
/** Stable application identity and Word revision addresses, without load-specific canonical node IDs. */
export interface PropertyReviewRecord {
  readonly schemaVersion: 1;
  readonly id: string;
  readonly author: string;
  readonly date: string;
  /** Pending proposals observed before this proposal. Missing, already-decided IDs do not create live owners. */
  readonly predecessors: readonly string[];
  readonly blocks: readonly StoredPropertyBlock[];
  readonly nativeRevisions: readonly StoredPropertyRevision[];
  readonly extension?: JsonObject;
}
export type PropertyRecordDecode =
  | { readonly status: 'supported'; readonly value: PropertyReviewRecord }
  | {
      readonly status: 'unsupported-version';
      readonly id: string;
      readonly version: number;
      readonly raw: string;
    }
  | { readonly status: 'invalid'; readonly reason: string };
export const PROPERTY_RECORD_LIMITS = Object.freeze({
  ...DEFAULT_PAYLOAD_LIMITS,
  maxBlocks: 1024,
  maxChangesPerBlock: 128,
  maxNativeRevisions: 4096,
});
const object = (value: unknown): value is JsonObject =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const keys = (value: object, names: readonly string[]) =>
  Object.keys(value).every((key) => names.includes(key));
const text = (value: unknown, limit: number): value is string =>
  typeof value === 'string' &&
  value.trim().length > 0 &&
  value.length <= limit &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) &&
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);
const state = (value: unknown): value is PlannedValue =>
  object(value) &&
  (value.present === false
    ? keys(value, ['present'])
    : value.present === true && Object.hasOwn(value, 'value') && keys(value, ['present', 'value']));
const equal = (a: PlannedValue, b: PlannedValue) =>
  a.present === b.present &&
  (!a.present || (b.present && jsonDataKey(a.value) === jsonDataKey(b.value)));

function schemaProblem(value: JsonValue): string | null {
  if (
    !object(value) ||
    value.schemaVersion !== 1 ||
    !keys(value, [
      'schemaVersion',
      'id',
      'author',
      'date',
      'blocks',
      'predecessors',
      'nativeRevisions',
      'extension',
    ])
  )
    return 'invalid-property-record-schema';
  if (!isBlockId(value.id)) return 'invalid-property-record-id';
  if (!text(value.author, 4096)) return 'invalid-property-record-author';
  if (
    typeof value.date !== 'string' ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value.date) ||
    !Number.isFinite(Date.parse(value.date)) ||
    new Date(value.date).toISOString() !== value.date
  )
    return 'invalid-property-record-date';
  if (
    !Array.isArray(value.predecessors) ||
    value.predecessors.length > 1024 ||
    value.predecessors.some((id) => !isBlockId(id) || id === value.id) ||
    new Set(value.predecessors).size !== value.predecessors.length
  )
    return 'invalid-property-record-predecessors';
  if (
    !Array.isArray(value.blocks) ||
    !value.blocks.length ||
    value.blocks.length > PROPERTY_RECORD_LIMITS.maxBlocks
  )
    return 'invalid-property-record-blocks';
  const blocks = new Set<string>();
  for (const block of value.blocks) {
    if (
      !object(block) ||
      !keys(block, ['blockId', 'type', 'changes']) ||
      !isBlockId(block.blockId) ||
      typeof block.type !== 'string' ||
      !/^[A-Za-z][A-Za-z0-9_.:-]{0,127}$/.test(block.type) ||
      blocks.has(block.blockId)
    )
      return 'invalid-or-duplicate-property-record-block';
    blocks.add(block.blockId);
    if (
      !Array.isArray(block.changes) ||
      !block.changes.length ||
      block.changes.length > PROPERTY_RECORD_LIMITS.maxChangesPerBlock
    )
      return 'invalid-property-record-changes';
    const names = new Set<string>();
    for (const change of block.changes) {
      if (
        !object(change) ||
        !keys(change, ['name', 'before', 'after']) ||
        !text(change.name, 128) ||
        names.has(change.name) ||
        !state(change.before) ||
        !state(change.after) ||
        equal(change.before, change.after)
      )
        return 'invalid-or-duplicate-property-record-change';
      names.add(change.name);
    }
  }
  if (
    !Array.isArray(value.nativeRevisions) ||
    value.nativeRevisions.length > PROPERTY_RECORD_LIMITS.maxNativeRevisions
  )
    return 'invalid-property-record-revisions';
  const revisions = new Set<string>();
  for (const revision of value.nativeRevisions) {
    if (
      !object(revision) ||
      !keys(revision, ['partName', 'id', 'author', 'date']) ||
      !text(revision.partName, 1024) ||
      !revision.partName.startsWith('/') ||
      revision.partName.includes('\\') ||
      revision.partName.split('/').some((segment) => segment === '..' || segment === '.') ||
      !text(revision.id, 256) ||
      !text(revision.author, 4096) ||
      (Object.hasOwn(revision, 'date') && !text(revision.date, 256))
    )
      return 'invalid-property-record-revision';
    const key = JSON.stringify([
      revision.partName,
      revision.id,
      revision.author,
      revision.date ?? null,
    ]);
    if (revisions.has(key)) return 'duplicate-property-record-revision';
    revisions.add(key);
  }
  if (Object.hasOwn(value, 'extension') && !object(value.extension))
    return 'invalid-property-record-extension';
  return null;
}
function freeze(value: JsonValue): void {
  if (value === null || typeof value !== 'object') return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}

/** No coercion or JavaScript evaluation. Future versions remain opaque, including their original bytes. */
export function decodePropertyReviewRecord(raw: string): PropertyRecordDecode {
  if (
    typeof raw !== 'string' ||
    raw.length > PROPERTY_RECORD_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > PROPERTY_RECORD_LIMITS.maxBytes
  )
    return { status: 'invalid', reason: 'byte-limit' };
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return { status: 'invalid', reason: 'invalid-json' };
  }
  const copied = copyBoundedJson(parsed, PROPERTY_RECORD_LIMITS);
  if (!copied.ok) return { status: 'invalid', reason: copied.reason };
  const value = copied.value;
  if (
    !object(value) ||
    !Number.isSafeInteger(value.schemaVersion) ||
    Number(value.schemaVersion) < 1 ||
    !isBlockId(value.id)
  )
    return { status: 'invalid', reason: 'invalid-property-record-version-or-id' };
  if (value.schemaVersion !== 1)
    return {
      status: 'unsupported-version',
      id: value.id,
      version: Number(value.schemaVersion),
      raw,
    };
  const problem = schemaProblem(value);
  if (problem) return { status: 'invalid', reason: problem };
  freeze(value);
  return { status: 'supported', value: value as unknown as PropertyReviewRecord };
}

/** Own the input before serialization can invoke a getter or discard a property. */
export function encodePropertyReviewRecord(record: PropertyReviewRecord): string {
  const copied = copyBoundedJson(record, PROPERTY_RECORD_LIMITS);
  if (!copied.ok) throw new TypeError(copied.reason);
  const problem = schemaProblem(copied.value);
  if (problem) throw new TypeError(problem);
  return JSON.stringify(copied.value);
}
