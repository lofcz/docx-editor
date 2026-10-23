import { decodeMentionPayload } from '../blocks/mention-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { SerializedBlockTag } from './block-tags';

/** Existing link-style mention syntax; escape label syntax without evaluating any content. */
export function serializeMention(payload: BlockPayload): SerializedBlockTag {
  const read = decodeMentionPayload(encodeBlockPayload(payload));
  if (read.status !== 'supported') throw new TypeError('invalid-mention-payload');
  const label = read.value.replace(/[\\`*_{}\[\]()<>&!~|$]|[\r\n\t]/g, (char) =>
    char === '\r' ? '&#13;' : char === '\n' ? '&#10;' : char === '\t' ? '&#9;' : '\\' + char
  );
  const identifier = encodeURIComponent(read.key).replace(
    /[!'()*]/g,
    (char) => '%' + char.charCodeAt(0).toString(16).toUpperCase()
  );
  const text = `[${label}](mention:${identifier})`;
  return {
    tag: 'mention',
    text,
    // Escapes/entities make raw value slices unreliable. Whole atom property ownership stays exact.
    properties: ['key', 'value'].map((property) => ({
      name: property,
      property,
      path: [property],
      start: 0,
      end: text.length,
      valueStart: null,
      valueEnd: null,
    })),
  };
}
