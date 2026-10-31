import type { MdxComponentRegistry } from '../registry/contract';
import { activeMdxRegistry } from '../registry/active';
import { isInlineComponentPayload } from '../registry/mdx-read';
import {
  decodeBlockPayload,
  encodeBlockPayload,
  type BlockPayload,
} from '../blocks/payload-schema';

/** Explicit native atom deletion. The executor verifies both payload and native subtree. */
export interface InlineRemoval {
  readonly kind: 'inline-remove';
  readonly paragraphId: string;
  readonly controlId: string;
  readonly blockId: string;
  readonly before: BlockPayload;
  readonly fingerprint: string;
}

export function removableInlinePayload(
  payload: BlockPayload,
  registry: MdxComponentRegistry = activeMdxRegistry()
): boolean {
  encodeBlockPayload(payload);
  return isInlineComponentPayload(payload, registry);
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
