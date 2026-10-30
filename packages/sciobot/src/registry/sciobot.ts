import { DOMAIN_BLOCK_TAGS, GENERATED_LABEL_TAG } from '../projection/block-tags';
import { domainDiffStrategy } from '../diff/domain-strategies';
import {
  CUSTOM_ALLOWED_TAGS,
  CUSTOM_DATA_ONLY_TAGS,
  CUSTOM_LITERAL_TAGS,
} from '../projection/mdx-tags';
import { BLOCK_TAG_PREFIX, PROPERTY_REVIEW_NODE_PREFIX } from '../blocks/payload-schema';
import { PAYLOAD_NAMESPACE, PAYLOAD_ROOT, PAYLOAD_NODE_PREFIX } from '../blocks/payload-opc-codec';
import {
  createMdxComponentRegistry,
  type MdxComponentDefinition,
  type MdxComponentRegistry,
} from './contract';

/** Payload types that sit inside a paragraph as one native atom. */
const INLINE_TYPES = new Set(['inline_equation', 'mention', 'inline_image', 'source']);
/** Payload types without MDX syntax of their own, kept as opaque application data. */
const OPAQUE_TYPES = ['listening'];

function definitions(): MdxComponentDefinition[] {
  const out: MdxComponentDefinition[] = [];
  for (const [type, mdx] of DOMAIN_BLOCK_TAGS) {
    const diff = domainDiffStrategy({ type });
    out.push({
      type,
      kind: INLINE_TYPES.has(type) ? 'inline' : 'block',
      content: mdx.kind === 'container' ? 'native' : 'none',
      mdx,
      ...(diff ? { diff } : {}),
    });
  }
  for (const type of OPAQUE_TYPES) {
    const diff = domainDiffStrategy({ type });
    out.push({ type, kind: 'block', content: 'opaque-mdx', ...(diff ? { diff } : {}) });
  }
  return out;
}

function build(): MdxComponentRegistry {
  const result = createMdxComponentRegistry({
    storage: {
      tagPrefix: BLOCK_TAG_PREFIX,
      namespace: PAYLOAD_NAMESPACE,
      root: PAYLOAD_ROOT,
      nodePrefix: PAYLOAD_NODE_PREFIX,
      reviewNodePrefix: PROPERTY_REVIEW_NODE_PREFIX,
      labelTag: GENERATED_LABEL_TAG,
    },
    components: definitions(),
    opaque: {
      allowed: CUSTOM_ALLOWED_TAGS,
      literal: CUSTOM_LITERAL_TAGS,
      dataOnly: CUSTOM_DATA_ONLY_TAGS,
    },
  });
  if (!result.ok) throw new Error(`Invalid Sciobot registry: ${result.reason} ${result.detail}`);
  return result.registry;
}

/**
 * The Sciobot components, storage names, and opaque MDX tags in one registry.
 *
 * It reads the existing adapter constants, so documents saved before the registry existed
 * keep their storage names. Later leaves move the definitions here and remove the constants.
 */
export const SCIOBOT_MDX_REGISTRY: MdxComponentRegistry = build();
