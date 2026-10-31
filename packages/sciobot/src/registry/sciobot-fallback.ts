import { decodeMathPayload, mathWordFallback } from '../blocks/math-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { MdxComponentDefinition } from './contract';
import { MdxRefusal } from './mdx-fields';

type WordFallback = NonNullable<MdxComponentDefinition['wordFallback']>;

/** A fenced code block shows its source lines in the inline code style. */
const code: WordFallback = (payload: BlockPayload) => {
  const language = payload.attributes.language,
    value = payload.attributes.value;
  if (typeof language !== 'string' || typeof value !== 'string' || language.length > 128)
    throw new MdxRefusal('A new code block needs bounded language and text values.');
  return { lines: value.split('\n'), code: true };
};

/** An equation shows its linear source, one paragraph per line. */
const equation: WordFallback = (payload: BlockPayload) => {
  const math = decodeMathPayload(encodeBlockPayload(payload));
  if (math.status !== 'supported')
    throw new MdxRefusal('A new equation needs a valid original source.');
  return { lines: mathWordFallback(math.payload).paragraphs };
};

export const SCIOBOT_WORD_FALLBACKS: Readonly<Record<string, WordFallback>> = {
  code_block: code,
  equation,
  inline_equation: equation,
};
