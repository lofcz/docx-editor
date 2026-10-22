import { isBlockId } from '../blocks/payload-schema';
import { copyBoundedJson, type JsonObject, type JsonValue } from '../data/json';
import { isReviewGroupRecordId } from './group-record-schema';

/** An immutable decision receipt, separate from its multi-block property proposal. */
export interface PayloadGroupDecisionRecord {
  readonly schemaVersion: 2;
  readonly id: string;
  readonly proposalId: string;
  readonly blockId: string;
  readonly decision: 'accept' | 'reject';
}
export const PAYLOAD_GROUP_DECISION_LIMITS = Object.freeze({
  maxBytes: 4096,
  maxDepth: 4,
  maxNodes: 32,
});
type DecisionDecode =
  | { readonly status: 'supported'; readonly value: PayloadGroupDecisionRecord }
  | { readonly status: 'invalid'; readonly reason: string };
const object = (value: unknown): value is JsonObject =>
  value !== null && typeof value === 'object' && !Array.isArray(value);

function problem(value: JsonValue): string | null {
  if (
    !object(value) ||
    value.schemaVersion !== 2 ||
    Object.keys(value).length !== 5 ||
    !Object.keys(value).every((key) =>
      ['schemaVersion', 'id', 'proposalId', 'blockId', 'decision'].includes(key)
    ) ||
    !isReviewGroupRecordId(value.id) ||
    !isBlockId(value.proposalId) ||
    !isBlockId(value.blockId) ||
    (value.decision !== 'accept' && value.decision !== 'reject')
  )
    return 'invalid-payload-group-decision';
  return null;
}

export function decodePayloadGroupDecisionRecord(raw: string): DecisionDecode {
  if (
    typeof raw !== 'string' ||
    raw.length > PAYLOAD_GROUP_DECISION_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > PAYLOAD_GROUP_DECISION_LIMITS.maxBytes
  )
    return { status: 'invalid', reason: 'byte-limit' };
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return { status: 'invalid', reason: 'invalid-json' };
  }
  const copied = copyBoundedJson(value, PAYLOAD_GROUP_DECISION_LIMITS);
  if (!copied.ok) return { status: 'invalid', reason: copied.reason };
  const issue = problem(copied.value);
  return issue
    ? { status: 'invalid', reason: issue }
    : {
        status: 'supported',
        value: Object.freeze(copied.value as unknown as PayloadGroupDecisionRecord),
      };
}

export function encodePayloadGroupDecisionRecord(record: PayloadGroupDecisionRecord): string {
  const copied = copyBoundedJson(record, PAYLOAD_GROUP_DECISION_LIMITS);
  if (!copied.ok) throw new TypeError(copied.reason);
  const issue = problem(copied.value);
  if (issue) throw new TypeError(issue);
  return JSON.stringify(copied.value);
}
