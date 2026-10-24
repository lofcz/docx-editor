import {
  decodeBlockPayload,
  encodeBlockPayload,
  type BlockPayload,
} from '../blocks/payload-schema';
import { decodeMathPayload } from '../blocks/math-codec';
import { decodeMentionPayload } from '../blocks/mention-codec';
import { decodeInlineImagePayload } from '../blocks/inline-image-codec';

/** Explicit native atom deletion. The executor verifies both payload and native subtree. */
export interface InlineRemoval {
  readonly kind: 'inline-remove';
  readonly paragraphId: string;
  readonly controlId: string;
  readonly blockId: string;
  readonly before: BlockPayload;
  readonly fingerprint: string;
}

export function removableInlinePayload(payload: BlockPayload): boolean {
  const raw = encodeBlockPayload(payload),
    math = decodeMathPayload(raw);
  return (
    (math.status === 'supported' && math.payload.type === 'inline_equation') ||
    decodeMentionPayload(raw).status === 'supported' ||
    !!decodeInlineImagePayload(raw) ||
    payload.type === 'source'
  );
}

export function isInlineRemoval(value: Record<string, unknown>): boolean {
  if (!value.before || typeof value.before !== 'object' || Array.isArray(value.before))
    return false;
  const decoded = decodeBlockPayload(JSON.stringify(value.before));
  return (
    decoded.status === 'supported' &&
    decoded.value.id === value.blockId &&
    removableInlinePayload(decoded.value) &&
    [value.paragraphId, value.controlId, value.blockId, value.fingerprint].every(
      (entry) => typeof entry === 'string' && entry.length > 0
    )
  );
}
