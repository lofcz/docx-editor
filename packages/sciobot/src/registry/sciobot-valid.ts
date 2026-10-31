import { decodeInlineImagePayload } from '../blocks/inline-image-codec';
import { decodeMathPayload } from '../blocks/math-codec';
import { decodeMentionPayload } from '../blocks/mention-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import { ISO_DATE } from './sciobot-inline';

/**
 * One SMILES or reaction SMILES string: printable ASCII without spaces, double quotes, `&`, `<`, or
 * braces, at most 4096 characters. It may be empty while the structure is not drawn yet.
 */
export const SMILES = /^[!#-%'-;=>?@-z|~]{0,4096}$/;

/** Payload checks for Sciobot components whose codecs constrain their attributes. */
export const SCIOBOT_VALIDATORS: Readonly<Record<string, (payload: BlockPayload) => boolean>> = {
  inline_equation: (payload) => {
    const math = decodeMathPayload(encodeBlockPayload(payload));
    return math.status === 'supported' && math.kind === 'inline';
  },
  equation: (payload) => {
    const math = decodeMathPayload(encodeBlockPayload(payload));
    return math.status === 'supported' && math.kind === 'block';
  },
  mention: (payload) => decodeMentionPayload(encodeBlockPayload(payload)).status === 'supported',
  inline_image: (payload) => !!decodeInlineImagePayload(encodeBlockPayload(payload)),
  smiles: (payload) =>
    Object.keys(payload.attributes).length === 1 &&
    typeof payload.attributes.smiles === 'string' &&
    SMILES.test(payload.attributes.smiles),
  date: (payload) =>
    Object.keys(payload.attributes).length === 1 &&
    typeof payload.attributes.value === 'string' &&
    ISO_DATE.test(payload.attributes.value),
};
