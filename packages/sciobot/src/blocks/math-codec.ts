import {
  decodeBlockPayload,
  encodeBlockPayload,
  type BlockPayload,
  type PayloadDecodeResult,
} from './payload-schema';

export const INLINE_MATH = 'inline_equation';
export const BLOCK_MATH = 'equation';
export type MathKind = 'inline' | 'block';
export interface MathPayload extends BlockPayload {
  readonly type: typeof INLINE_MATH | typeof BLOCK_MATH;
  readonly attributes: BlockPayload['attributes'] & { readonly texExpression: string };
}
export type MathPayloadRead =
  | {
      readonly status: 'supported';
      readonly kind: MathKind;
      readonly source: string;
      readonly payload: MathPayload;
    }
  | Exclude<PayloadDecodeResult, { status: 'supported' }>;
const sourceProblem = (value: unknown): string | null => {
  if (typeof value !== 'string') return 'invalid-math-source';
  if (
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) ||
    /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value)
  )
    return 'invalid-math-source-character';
  return null;
};
/** Unknown versions and unrelated payload fields remain opaque. No LaTeX is evaluated. */
export function decodeMathPayload(raw: string): MathPayloadRead {
  const read = decodeBlockPayload(raw);
  if (read.status !== 'supported') return read;
  const kind =
    read.value.type === INLINE_MATH ? 'inline' : read.value.type === BLOCK_MATH ? 'block' : null;
  const problem = kind ? sourceProblem(read.value.attributes.texExpression) : 'not-math-payload';
  if (problem) return { status: 'invalid', reason: problem, raw };
  return {
    status: 'supported',
    kind: kind!,
    source: read.value.attributes.texExpression as string,
    payload: read.value as MathPayload,
  };
}
/** Construct a bounded application equation, retaining source whitespace and Unicode exactly. */
export function createMathPayload(id: string, kind: MathKind, texExpression: string): MathPayload {
  if (kind !== 'inline' && kind !== 'block') throw new TypeError('invalid-math-kind');
  const problem = sourceProblem(texExpression);
  if (problem) throw new TypeError(problem);
  const payload: MathPayload = {
    version: 1,
    id,
    type: kind === 'inline' ? INLINE_MATH : BLOCK_MATH,
    attributes: { texExpression },
  };
  return JSON.parse(encodeBlockPayload(payload)) as MathPayload;
}
/** Change only the source, without discarding private attributes or future payload fields. */
export function withMathSource(payload: BlockPayload, texExpression: string): MathPayload {
  const read = decodeMathPayload(encodeBlockPayload(payload));
  if (read.status !== 'supported')
    throw new TypeError(read.status === 'invalid' ? read.reason : 'unsupported-math-version');
  const problem = sourceProblem(texExpression);
  if (problem) throw new TypeError(problem);
  return JSON.parse(
    encodeBlockPayload({
      ...read.payload,
      attributes: { ...read.payload.attributes, texExpression },
    })
  ) as MathPayload;
}
/** Word receives readable source text; the authoritative payload never normalizes line endings. */
export function mathWordFallback(payload: MathPayload): {
  readonly label: string;
  readonly inlineText: string;
  readonly paragraphs: readonly string[];
} {
  const read = decodeMathPayload(encodeBlockPayload(payload));
  if (read.status !== 'supported') throw new TypeError('invalid-math-payload');
  const lines = read.source.split(/\r\n|\r|\n/);
  const paragraphs = read.source.length ? lines : ['□'];
  return {
    label: Array.from(lines.find((line) => line.length > 0) ?? '□')
      .slice(0, 128)
      .join(''),
    inlineText: paragraphs.join(' '),
    paragraphs,
  };
}
