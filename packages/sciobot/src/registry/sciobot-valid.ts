import { decodeInlineImagePayload } from '../blocks/inline-image-codec';
import { decodeMathPayload } from '../blocks/math-codec';
import { decodeMentionPayload } from '../blocks/mention-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';

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
};
