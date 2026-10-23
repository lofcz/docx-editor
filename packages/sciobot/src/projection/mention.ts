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
    // Field ranges distinguish label syntax from identifier syntax. Escapes have no direct value slice.
    properties: ['key', 'value'].map((property) => ({
      name: property,
      property,
      path: [property],
      start: property === 'key' ? label.length + 2 : 0,
      end: property === 'key' ? text.length : label.length + 2,
      valueStart: null,
      valueEnd: null,
    })),
  };
}
