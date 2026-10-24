import {
  MAX_REVIEW_FIELD_BYTES,
  MAX_REVIEW_FIELD_INTENTS,
  decodeFieldIntent,
} from './review-field-codec';
import { MAX_REVIEW_FLAG_INTENTS, validateFlagIntent } from './review-flag-codec';
import { REVIEW_GROUP_STORE_LIMITS } from '../review/group-record-limits';
import { GROUP_ID_LIMIT } from '../review/group-markers';

export const PROPERTY_GROUP_RECEIPTS = 'sciobot-property-group-receipts-v1';
export const MAX_PROPERTY_GROUPS = REVIEW_GROUP_STORE_LIMITS.maxRecords;
export interface PropertyGroupBinding {
  readonly schemaVersion: 1;
  readonly groupId: string;
  readonly membersKey: string;
  readonly fields: readonly string[];
  readonly flags: readonly string[];
}
const bytes = (value: string) => new TextEncoder().encode(value).length;
export function decodePropertyGroupBinding(key: string, raw: unknown): PropertyGroupBinding {
  if (typeof raw !== 'string' || bytes(raw) > MAX_REVIEW_FIELD_BYTES)
    throw new Error('Invalid bounded property group binding.');
  const value = JSON.parse(raw);
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).length !== 5 ||
    value.schemaVersion !== 1 ||
    value.groupId !== key ||
    typeof key !== 'string' ||
    !/^g1:stored:rg_[A-Za-z0-9_-]+$/.test(key) ||
    key.length > GROUP_ID_LIMIT ||
    typeof value.membersKey !== 'string' ||
    !value.membersKey.length ||
    !Array.isArray(value.fields) ||
    !Array.isArray(value.flags) ||
    value.fields.length > MAX_REVIEW_FIELD_INTENTS ||
    value.flags.length > MAX_REVIEW_FLAG_INTENTS
  )
    throw new Error('Invalid property group binding fields.');
  for (const [ids, flags] of [
    [value.fields, false],
    [value.flags, true],
  ] as const) {
    const known = new Set<string>();
    for (const id of ids) {
      if (typeof id !== 'string' || known.has(id))
        throw new Error('Duplicate property group intent.');
      if (flags) validateFlagIntent(id, 'target');
      else
        decodeFieldIntent(id, JSON.stringify({ target: 'target', value: null, stamp: [0, 0, 0] }));
      known.add(id);
    }
  }
  if (!value.fields.length && !value.flags.length) throw new Error('Empty property group binding.');
  return Object.freeze({
    ...value,
    fields: Object.freeze([...value.fields]),
    flags: Object.freeze([...value.flags]),
  });
}
export function encodePropertyGroupBinding(value: PropertyGroupBinding): string {
  const raw = JSON.stringify(value);
  decodePropertyGroupBinding(value.groupId, raw);
  return raw;
}
