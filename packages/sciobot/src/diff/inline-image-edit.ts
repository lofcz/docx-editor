import { decodeInlineImagePayload, type InlineImagePayload } from '../blocks/inline-image-codec';
import type { InsertedInlineAtom } from './inserted-atoms';
export type InlineImageResource = Extract<InsertedInlineAtom, { kind: 'drawing' }>['resource'];
/** Replace a managed image resource in place. Geometry and unmapped native metadata stay native. */
export interface InlineImageEdit {
  readonly kind: 'inline-image';
  readonly paragraphId: string;
  readonly controlId: string;
  readonly blockId: string;
  readonly drawingNodeId: string;
  readonly before: InlineImagePayload;
  readonly resource: InlineImageResource;
}
const record = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const text = (value: unknown): value is string =>
  typeof value === 'string' &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) &&
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);
export function isInlineImageEdit(value: Record<string, unknown>): boolean {
  const before = record(value.before) && decodeInlineImagePayload(JSON.stringify(value.before));
  const resource = value.resource;
  return (
    !!before &&
    before.id === value.blockId &&
    [value.paragraphId, value.controlId, value.blockId, value.drawingNodeId].every(
      (id) => typeof id === 'string' && id.length > 0
    ) &&
    record(resource) &&
    text(resource.reference) &&
    resource.reference.length > 0 &&
    text(resource.alt) &&
    (resource.title === null || text(resource.title)) &&
    Object.keys(resource).every((key) => ['reference', 'alt', 'title'].includes(key))
  );
}
