import { parentNodeOf, type OoxmlPart } from '@docx-editor.dev/core/store';
import { serializeLinkedMention } from './mention-literal';
import { decodeMentionPayload } from '../blocks/mention-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { SerializedBlockTag } from './block-tags';

/** Existing link-style mention syntax; escape label syntax without evaluating any content. */
export function serializeMention(
  payload: BlockPayload,
  insideHyperlink = false
): SerializedBlockTag {
  const read = decodeMentionPayload(encodeBlockPayload(payload));
  if (read.status !== 'supported') throw new TypeError('invalid-mention-payload');
  if (insideHyperlink) return serializeLinkedMention(read.key, read.value);
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

/** Revision and extension wrappers do not hide an enclosing hyperlink from mention syntax. */
export function mentionInsideHyperlink(part: OoxmlPart, controlId: string): boolean {
  let parent = parentNodeOf(part, controlId);
  for (let depth = 0; parent && depth < 64; depth++) {
    if (parent.kind === 'hyperlink') return true;
    if (parent.kind === 'paragraph') return false;
    parent = parentNodeOf(part, parent.id);
  }
  return false;
}
