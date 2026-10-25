import type { BlockPayload } from '../blocks/payload-schema';
import type { SerializedBlockTag } from './block-tags';

const fenceFor = (value: string): string => {
  let longest = 0;
  for (const match of value.matchAll(/`+/g)) longest = Math.max(longest, match[0]!.length);
  return '`'.repeat(Math.max(3, longest + 1));
};

/** Preserve a fenced code block's language and exact body without interpreting its contents. */
export function serializeCodeBlock(payload: BlockPayload): SerializedBlockTag {
  const language = payload.attributes.language;
  const value = payload.attributes.value;
  if (typeof language !== 'string' || typeof value !== 'string' || language.length > 128)
    throw new TypeError('invalid-code-block-payload');
  const fence = fenceFor(value);
  const text = `${fence}${language ?? ''}\n${value}\n${fence}`;
  return {
    tag: 'code',
    text,
    properties: ['language', 'value'].map((property) => ({
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
