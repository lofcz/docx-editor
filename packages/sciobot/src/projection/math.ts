import { normalizeTargetSource } from './normalize-target';
import remarkMath from 'remark-math';
import remarkParse from 'remark-parse';
import remarkStringify from 'remark-stringify';
import { unified } from 'unified';
import { BLOCK_MATH, INLINE_MATH, decodeMathPayload } from '../blocks/math-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { SerializedBlockTag } from './block-tags';
const processor = unified().use(remarkParse).use(remarkMath).use(remarkStringify);

function unescapedPipe(text: string): boolean {
  let slashes = 0;
  for (const char of text) {
    if (char === '|' && slashes % 2 === 0) return true;
    slashes = char === '\\' ? slashes + 1 : 0;
  }
  return false;
}
/** Application-reserved tags supply lossless data syntax when math fences cannot. */
function fallback(type: typeof INLINE_MATH | typeof BLOCK_MATH, source: string): string {
  const escaped = source.replace(
    /[&"<>\r\n\t|]/g,
    (char) =>
      ({
        '&': '&amp;',
        '"': '&quot;',
        '<': '&lt;',
        '>': '&gt;',
        '\r': '&#13;',
        '\n': '&#10;',
        '\t': '&#9;',
        '|': '&#124;',
      })[char]!
  );
  const tag = type === INLINE_MATH ? 'scio-inline-equation' : 'scio-equation';
  return `<${tag} texExpression="${escaped}" />`;
}
/** Emit standard math only when it parses back to the exact original source and kind. */
export function serializeMath(payload: BlockPayload): SerializedBlockTag {
  const read = decodeMathPayload(encodeBlockPayload(payload));
  if (read.status !== 'supported') throw new TypeError('invalid-math-payload');
  const type = read.kind === 'inline' ? 'inlineMath' : 'math';
  const node = { type, value: read.source };
  const root = {
    type: 'root',
    children: read.kind === 'inline' ? [{ type: 'paragraph', children: [node] }] : [node],
  };
  const printed = processor.stringify(root as Parameters<typeof processor.stringify>[0]);
  const candidate = printed.endsWith('\n') ? printed.slice(0, -1) : printed;
  const parsed = processor.parse(normalizeTargetSource(candidate).markdown);
  const children = parsed.children;
  const recovered =
    children.length === 1 && children[0]!.type === 'paragraph'
      ? children[0]!.children.length === 1
        ? children[0]!.children[0]
        : null
      : children.length === 1
        ? children[0]
        : null;
  const exact = recovered?.type === type && 'value' in recovered && recovered.value === read.source;
  const text =
    exact && (read.kind === 'block' || !unescapedPipe(candidate))
      ? candidate
      : fallback(read.payload.type, read.source);
  return {
    tag: read.kind === 'inline' ? 'scio-inline-equation' : 'scio-equation',
    text,
    properties: [
      {
        name: 'texExpression',
        property: 'texExpression',
        path: ['texExpression'],
        start: 0,
        end: text.length,
        valueStart: null,
        valueEnd: null,
      },
    ],
  };
}
