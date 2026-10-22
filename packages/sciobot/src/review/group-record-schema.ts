import type { RevisionAddress } from '@docx-editor.dev/core/store';
import { copyBoundedJson, type JsonObject, type JsonValue } from '../data/json';

export interface StoredReviewGroupMember {
  readonly partName: string;
  readonly revision: RevisionAddress;
}
/** Native Word addresses only. Canonical node IDs and XML paths belong to live reads. */
export interface ReviewGroupRecord {
  readonly schemaVersion: 1;
  readonly id: string;
  readonly members: readonly StoredReviewGroupMember[];
  readonly extension?: JsonObject;
}
export type ReviewGroupDecode =
  | { readonly status: 'supported'; readonly value: ReviewGroupRecord }
  | {
      readonly status: 'unsupported-version';
      readonly id: string;
      readonly version: number;
      readonly raw: string;
    }
  | { readonly status: 'invalid'; readonly reason: string };
export const REVIEW_GROUP_RECORD_LIMITS = Object.freeze({
  maxBytes: 2 * 1024 * 1024,
  maxDepth: 32,
  maxNodes: 50_000,
  maxMembers: 4096,
});
export const isReviewGroupRecordId = (id: unknown): id is string =>
  typeof id === 'string' && /^rg_[A-Za-z0-9_-]{1,96}$/.test(id);
export const reviewGroupMemberKey = (member: StoredReviewGroupMember) =>
  JSON.stringify([
    member.partName,
    member.revision.id,
    member.revision.author,
    member.revision.date ?? null,
  ]);
const object = (value: unknown): value is JsonObject =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const keys = (value: object, names: readonly string[]) =>
  Object.keys(value).every((name) => names.includes(name));
const text = (value: unknown, limit: number, empty = false): value is string =>
  typeof value === 'string' &&
  (empty || value.trim().length > 0) &&
  value.length <= limit &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) &&
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);

function problem(value: JsonValue): string | null {
  if (
    !object(value) ||
    value.schemaVersion !== 1 ||
    !keys(value, ['schemaVersion', 'id', 'members', 'extension'])
  )
    return 'invalid-review-group-schema';
  if (!isReviewGroupRecordId(value.id)) return 'invalid-review-group-id';
  if (
    !Array.isArray(value.members) ||
    !value.members.length ||
    value.members.length > REVIEW_GROUP_RECORD_LIMITS.maxMembers
  )
    return 'invalid-review-group-members';
  const members = new Set<string>();
  for (const member of value.members) {
    if (
      !object(member) ||
      !keys(member, ['partName', 'revision']) ||
      !text(member.partName, 1024) ||
      !member.partName.startsWith('/') ||
      member.partName.includes('\\') ||
      member.partName.split('/').some((part) => part === '..' || part === '.') ||
      !object(member.revision) ||
      !keys(member.revision, ['id', 'author', 'date']) ||
      !text(member.revision.id, 256) ||
      !text(member.revision.author, 4096, true) ||
      (Object.hasOwn(member.revision, 'date') && !text(member.revision.date, 256))
    )
      return 'invalid-review-group-member';
    const key = reviewGroupMemberKey(member as unknown as StoredReviewGroupMember);
    if (members.has(key)) return 'duplicate-review-group-member';
    members.add(key);
  }
  if (Object.hasOwn(value, 'extension') && !object(value.extension))
    return 'invalid-review-group-extension';
  return null;
}
function freeze(value: JsonValue): void {
  if (value === null || typeof value !== 'object') return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}
export function decodeReviewGroupRecord(raw: string): ReviewGroupDecode {
  if (
    typeof raw !== 'string' ||
    raw.length > REVIEW_GROUP_RECORD_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > REVIEW_GROUP_RECORD_LIMITS.maxBytes
  )
    return { status: 'invalid', reason: 'byte-limit' };
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return { status: 'invalid', reason: 'invalid-json' };
  }
  const copied = copyBoundedJson(value, REVIEW_GROUP_RECORD_LIMITS);
  if (!copied.ok) return { status: 'invalid', reason: copied.reason };
  const data = copied.value;
  if (
    object(data) &&
    Number.isSafeInteger(data.schemaVersion) &&
    Number(data.schemaVersion) > 1 &&
    isReviewGroupRecordId(data.id)
  )
    return { status: 'unsupported-version', id: data.id, version: Number(data.schemaVersion), raw };
  const issue = problem(data);
  if (issue) return { status: 'invalid', reason: issue };
  freeze(data);
  return { status: 'supported', value: data as unknown as ReviewGroupRecord };
}
export function encodeReviewGroupRecord(record: ReviewGroupRecord): string {
  const copied = copyBoundedJson(record, REVIEW_GROUP_RECORD_LIMITS);
  if (!copied.ok) throw new TypeError(copied.reason);
  const issue = problem(copied.value);
  if (issue) throw new TypeError(issue);
  return JSON.stringify(copied.value);
}
