import { decodeMathPayload, mathWordFallback } from '../blocks/math-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { MdxComponentDefinition } from './contract';
import { MdxRefusal } from './mdx-fields';
import { SMILES } from './sciobot-valid';

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

/** A table of contents is derived from the headings; Word shows its title line. */
const toc: WordFallback = () => ({ lines: ['Table of contents'] });

/** A SMILES structure shows its exact source in the inline code style. */
const smiles: WordFallback = (payload: BlockPayload) => {
  if (!SMILES.test(String(payload.attributes.smiles)))
    throw new MdxRefusal('A SMILES block needs a bounded single-line SMILES string.');
  return { lines: [payload.attributes.smiles as string], code: true };
};

export const SCIOBOT_WORD_FALLBACKS: Readonly<Record<string, WordFallback>> = {
  code_block: code,
  toc,
  smiles,
  equation,
  inline_equation: equation,
};
