import {
  decodeBlockPayload,
  encodeBlockPayload,
  type BlockPayload,
  type PayloadDecodeResult,
} from './payload-schema';
export const MENTION = 'mention';
export interface MentionPayload extends BlockPayload {
  readonly type: typeof MENTION;
  readonly attributes: BlockPayload['attributes'] & {
    readonly key: string;
    readonly value: string;
  };
}
export type MentionPayloadRead =
  | {
      readonly status: 'supported';
      readonly payload: MentionPayload;
      readonly key: string;
      readonly value: string;
    }
  | Exclude<PayloadDecodeResult, { status: 'supported' }>;
const validText = (value: unknown): value is string =>
  typeof value === 'string' &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) &&
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);
/** Person/document identity is application metadata, independent from the persistent block ID. */
export function decodeMentionPayload(raw: string): MentionPayloadRead {
  const read = decodeBlockPayload(raw);
  if (read.status !== 'supported') return read;
  if (read.value.type !== MENTION) return { status: 'invalid', reason: 'not-mention-payload', raw };
  const { key, value } = read.value.attributes;
  if (!validText(key) || !key.length || !validText(value))
    return { status: 'invalid', reason: 'invalid-mention-identity', raw };
  return { status: 'supported', key, value, payload: read.value as MentionPayload };
}
export function createMentionPayload(
  id: string,
  key: string,
  value: string,
  attributes: BlockPayload['attributes'] = {}
): MentionPayload {
  // Validate unknown attributes before copying. Accessors and executable values never run.
  const input = encodeBlockPayload({ id, type: MENTION, version: 1, attributes });
  const base = JSON.parse(input) as BlockPayload;
  const raw = encodeBlockPayload({ ...base, attributes: { ...base.attributes, key, value } });
  const read = decodeMentionPayload(raw);
  if (read.status !== 'supported')
    throw new TypeError(read.status === 'invalid' ? read.reason : 'unsupported-mention-version');
  return read.payload;
}
/** Word fallback normalizes line boundaries only; stored display text stays exact. */
export function mentionWordFallback(payload: MentionPayload): {
  readonly label: string;
  readonly text: string;
} {
  const read = decodeMentionPayload(encodeBlockPayload(payload));
  if (read.status !== 'supported') throw new TypeError('invalid-mention-payload');
  const text = (read.value || read.key).replace(/\r\n|\r|\n|\t/g, ' ');
  return { text, label: Array.from(text).slice(0, 128).join('') };
}
