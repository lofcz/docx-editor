import remarkParse from 'remark-parse';
import remarkStringify from 'remark-stringify';
import { unified } from 'unified';
import { decodeInlineImagePayload } from '../blocks/inline-image-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import { normalizeTargetSource } from './normalize-target';
import type { SerializedBlockTag } from './block-tags';
const processor = unified().use(remarkParse).use(remarkStringify);
const attribute = (value: string) =>
  value.replace(
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
export function serializeInlineImage(payload: BlockPayload): SerializedBlockTag {
  const read = decodeInlineImagePayload(encodeBlockPayload(payload));
  if (!read) throw new TypeError('invalid-inline-image-payload');
  const a = read.attributes;
  const image = { type: 'image', url: a.reference, alt: a.alt, title: a.title };
  const candidate = processor
    .stringify({ type: 'root', children: [{ type: 'paragraph', children: [image] }] } as Parameters<
      typeof processor.stringify
    >[0])
    .trimEnd();
  const parsed = processor.parse(normalizeTargetSource('Before ' + candidate + ' after').markdown);
  const recovered =
    parsed.children[0]?.type === 'paragraph'
      ? parsed.children[0].children.find((node) => node.type === 'image')
      : null;
  const exact =
    recovered?.type === 'image' &&
    recovered.url === a.reference &&
    recovered.alt === a.alt &&
    recovered.title === a.title;
  let slashes = 0,
    pipe = false;
  for (const char of candidate) {
    if (char === '|' && slashes % 2 === 0) pipe = true;
    slashes = char === '\\' ? slashes + 1 : 0;
  }
  const text =
    exact && !pipe
      ? candidate
      : `<scio-inline-image reference="${attribute(a.reference)}" alt="${attribute(a.alt)}"${a.title !== null ? ` title="${attribute(a.title)}"` : ''} />`;
  return { tag: 'scio-inline-image', text, properties: [] };
}
