import { decodeBlockPayload, encodeBlockPayload, type BlockPayload } from './payload-schema';
export const INLINE_IMAGE = 'inline_image';
export interface InlineImagePayload extends BlockPayload {
  readonly type: typeof INLINE_IMAGE;
  readonly attributes: BlockPayload['attributes'] & {
    readonly reference: string;
    readonly alt: string;
    readonly title: string | null;
    readonly contentHash: string;
    readonly mediaPartName: string;
    readonly widthPoints: number;
    readonly heightPoints: number;
  };
}
const text = (value: unknown): value is string =>
  typeof value === 'string' &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) &&
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);
export function decodeInlineImagePayload(raw: string): InlineImagePayload | null {
  const read = decodeBlockPayload(raw);
  if (read.status !== 'supported' || read.value.type !== INLINE_IMAGE) return null;
  const a = read.value.attributes;
  return text(a.reference) &&
    !!a.reference &&
    text(a.alt) &&
    (a.title === null || text(a.title)) &&
    typeof a.contentHash === 'string' &&
    /^[a-f0-9]{64}$/.test(a.contentHash) &&
    text(a.mediaPartName) &&
    a.mediaPartName.startsWith('/') &&
    typeof a.widthPoints === 'number' &&
    Number.isFinite(a.widthPoints) &&
    a.widthPoints > 0 &&
    typeof a.heightPoints === 'number' &&
    Number.isFinite(a.heightPoints) &&
    a.heightPoints > 0
    ? (read.value as InlineImagePayload)
    : null;
}
export function createInlineImagePayload(
  id: string,
  attributes: InlineImagePayload['attributes']
): InlineImagePayload {
  const payload: InlineImagePayload = { version: 1, id, type: INLINE_IMAGE, attributes };
  if (!decodeInlineImagePayload(encodeBlockPayload(payload)))
    throw new TypeError('invalid-inline-image-payload');
  return payload;
}
