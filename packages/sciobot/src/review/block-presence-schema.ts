import { normalizePartName } from '@docx-editor.dev/core/store';
import {
  DEFAULT_PAYLOAD_LIMITS,
  decodeBlockPayload,
  isBlockId,
  type BlockPayload,
} from '../blocks/payload-schema';
import { isPayloadLabel } from '../blocks/payload-label';
import { copyBoundedJson, type JsonObject, type JsonValue } from '../data/json';

export const BLOCK_PRESENCE_EXECUTOR = 'sciobot-block-presence-v1';
export interface BlockPresenceMember extends JsonObject {
  readonly payload: BlockPayload;
  readonly label: string;
}
export interface BlockPresenceChange {
  readonly blockId: string;
  readonly partName: string;
  readonly before: { readonly present: boolean };
  readonly after: { readonly present: boolean };
  /** Every application payload in this control subtree, including its root. */
  readonly members: readonly BlockPresenceMember[];
  /** Native ownership proof, without copied Word content or canonical node IDs. */
  readonly nativeFingerprint: `sha256:${string}`;
}
/** Private profiles own blocks, retained atoms, paragraph births, cell births, and whole sibling gaps. */
export interface BlockPresenceRecord {
  readonly schemaVersion: 2 | 3 | 4 | 5 | 6;
  readonly executor: typeof BLOCK_PRESENCE_EXECUTOR;
  readonly id: string;
  readonly author: string;
  readonly date: string;
  readonly predecessors: readonly string[];
  readonly blocks: readonly BlockPresenceChange[];
  readonly extension?: JsonObject;
}
export type BlockPresenceDecode =
  | { readonly status: 'supported'; readonly value: BlockPresenceRecord }
  | {
      readonly status: 'unsupported-version';
      readonly id: string;
      readonly version: number;
      readonly raw: string;
    }
  | { readonly status: 'invalid'; readonly reason: string };
export const BLOCK_PRESENCE_LIMITS = Object.freeze({
  ...DEFAULT_PAYLOAD_LIMITS,
  maxBlocks: 1024,
  maxMembers: 1024,
});
const object = (value: unknown): value is JsonObject =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const keys = (value: object, allowed: readonly string[]) =>
  Object.keys(value).every((key) => allowed.includes(key));
const plainText = (value: unknown, max: number): value is string =>
  typeof value === 'string' &&
  value.trim().length > 0 &&
  value.length <= max &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) &&
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);
export function blockPresenceAttributionProblem(
  id: unknown,
  author: unknown,
  date: unknown
): string | null {
  if (!isBlockId(id) || !plainText(author, 4096)) return 'invalid-presence-attribution';
  if (
    typeof date !== 'string' ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(date) ||
    !Number.isFinite(Date.parse(date)) ||
    new Date(date).toISOString() !== date
  )
    return 'invalid-presence-date';
  return null;
}
function problem(value: JsonValue): string | null {
  if (
    !object(value) ||
    (value.schemaVersion !== 2 &&
      value.schemaVersion !== 3 &&
      value.schemaVersion !== 4 &&
      value.schemaVersion !== 5 &&
      value.schemaVersion !== 6) ||
    value.executor !== BLOCK_PRESENCE_EXECUTOR ||
    !keys(value, [
      'schemaVersion',
      'executor',
      'id',
      'author',
      'date',
      'predecessors',
      'blocks',
      'extension',
    ])
  )
    return 'invalid-presence-schema';
  const inline = object(value.extension) && Object.hasOwn(value.extension, 'nativeInline');
  if (
    (value.schemaVersion === 3 ||
      value.schemaVersion === 4 ||
      value.schemaVersion === 5 ||
      value.schemaVersion === 6) !== inline
  )
    return 'invalid-inline-presence-version';
  const attribution = blockPresenceAttributionProblem(value.id, value.author, value.date);
  if (attribution) return attribution;
  if (
    !Array.isArray(value.predecessors) ||
    value.predecessors.length > BLOCK_PRESENCE_LIMITS.maxBlocks ||
    value.predecessors.some((id) => !isBlockId(id) || id === value.id) ||
    new Set(value.predecessors).size !== value.predecessors.length
  )
    return 'invalid-presence-predecessors';
  if (
    !Array.isArray(value.blocks) ||
    !value.blocks.length ||
    value.blocks.length > BLOCK_PRESENCE_LIMITS.maxBlocks
  )
    return 'invalid-presence-blocks';
  const members = new Set<string>(),
    roots = new Set<string>();
  for (const block of value.blocks) {
    if (
      !object(block) ||
      !keys(block, ['blockId', 'partName', 'before', 'after', 'members', 'nativeFingerprint']) ||
      !isBlockId(block.blockId) ||
      roots.has(block.blockId)
    )
      return 'invalid-presence-root';
    roots.add(block.blockId);
    if (
      typeof block.partName !== 'string' ||
      block.partName.length > 1024 ||
      !block.partName.startsWith('/')
    )
      return 'invalid-presence-part';
    const part = normalizePartName(block.partName);
    if (!part.ok || part.partName !== block.partName) return 'invalid-presence-part';
    if (
      !object(block.before) ||
      !object(block.after) ||
      !keys(block.before, ['present']) ||
      !keys(block.after, ['present']) ||
      typeof block.before.present !== 'boolean' ||
      typeof block.after.present !== 'boolean' ||
      block.before.present === block.after.present
    )
      return 'invalid-presence-transition';
    if (
      typeof block.nativeFingerprint !== 'string' ||
      !/^sha256:[a-f0-9]{64}$/.test(block.nativeFingerprint)
    )
      return 'invalid-presence-native-proof';
    if (
      !Array.isArray(block.members) ||
      !block.members.length ||
      block.members.length > BLOCK_PRESENCE_LIMITS.maxMembers
    )
      return 'invalid-presence-members';
    let root = false;
    for (const member of block.members) {
      if (
        !object(member) ||
        !keys(member, ['payload', 'label']) ||
        !object(member.payload) ||
        !isPayloadLabel(member.label)
      )
        return 'invalid-presence-member';
      const payload = decodeBlockPayload(JSON.stringify(member.payload));
      if (payload.status !== 'supported' || members.has(payload.value.id))
        return 'invalid-or-duplicate-presence-payload';
      members.add(payload.value.id);
      if (members.size > BLOCK_PRESENCE_LIMITS.maxMembers) return 'presence-member-limit';
      root ||= payload.value.id === block.blockId;
    }
    if (!root) return 'missing-presence-root-payload';
  }
  if (Object.hasOwn(value, 'extension') && !object(value.extension))
    return 'invalid-presence-extension';
  return null;
}
function freeze(value: JsonValue): void {
  if (value === null || typeof value !== 'object') return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}
export function decodeBlockPresenceRecord(raw: string): BlockPresenceDecode {
  if (
    typeof raw !== 'string' ||
    raw.length > BLOCK_PRESENCE_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > BLOCK_PRESENCE_LIMITS.maxBytes
  )
    return { status: 'invalid', reason: 'byte-limit' };
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return { status: 'invalid', reason: 'invalid-json' };
  }
  const copied = copyBoundedJson(parsed, BLOCK_PRESENCE_LIMITS);
  if (!copied.ok) return { status: 'invalid', reason: copied.reason };
  const value = copied.value;
  if (
    !object(value) ||
    !isBlockId(value.id) ||
    !Number.isSafeInteger(value.schemaVersion) ||
    Number(value.schemaVersion) < 1
  )
    return { status: 'invalid', reason: 'invalid-presence-version-or-id' };
  if (
    value.schemaVersion !== 2 &&
    !(
      (value.schemaVersion === 3 ||
        value.schemaVersion === 4 ||
        value.schemaVersion === 5 ||
        value.schemaVersion === 6) &&
      object(value.extension) &&
      Object.hasOwn(value.extension, 'nativeInline')
    )
  )
    return {
      status: 'unsupported-version',
      id: value.id,
      version: Number(value.schemaVersion),
      raw,
    };
  const reason = problem(value);
  if (reason) return { status: 'invalid', reason };
  freeze(value);
  return { status: 'supported', value: value as unknown as BlockPresenceRecord };
}
export function encodeBlockPresenceRecord(record: BlockPresenceRecord): string {
  const copied = copyBoundedJson(record, BLOCK_PRESENCE_LIMITS);
  if (!copied.ok) throw new TypeError(copied.reason);
  const reason = problem(copied.value);
  if (reason) throw new TypeError(reason);
  return JSON.stringify(copied.value);
}
