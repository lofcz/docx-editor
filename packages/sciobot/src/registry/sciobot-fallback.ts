import { decodeMathPayload, mathWordFallback } from '../blocks/math-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { MdxComponentDefinition } from './contract';
import { MdxRefusal } from './mdx-fields';
import { SMILES } from './sciobot-valid';
import { validGallery } from '../projection/gallery';
import type { JsonObject } from '../data/json';

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

/** A gallery shows one line per image: its caption, or its address. */
const gallery: WordFallback = (payload: BlockPayload) => {
  if (!validGallery(payload))
    throw new MdxRefusal('A gallery needs one to 100 images with http(s) addresses.');
  return {
    lines: (payload.attributes.images as JsonObject[]).map(
      (image) => (image.caption as string | undefined) || (image.url as string)
    ),
  };
};

export const SCIOBOT_WORD_FALLBACKS: Readonly<Record<string, WordFallback>> = {
  code_block: code,
  toc,
  smiles,
  image_gallery: gallery,
  equation,
  inline_equation: equation,
};
