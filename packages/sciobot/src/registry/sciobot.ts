import { DOMAIN_BLOCK_TAGS, GENERATED_LABEL_TAG } from '../projection/block-tags';
import { SCIOBOT_DIFF_STRATEGIES } from './sciobot-diff';
import {
  CUSTOM_ALLOWED_TAGS,
  CUSTOM_DATA_ONLY_TAGS,
  CUSTOM_LITERAL_TAGS,
} from '../projection/mdx-tags';
import { BLOCK_TAG_PREFIX, PROPERTY_REVIEW_NODE_PREFIX } from '../blocks/payload-schema';
import { PAYLOAD_NAMESPACE, PAYLOAD_ROOT, PAYLOAD_NODE_PREFIX } from '../blocks/payload-opc-codec';
import { SCIOBOT_PARSERS } from './sciobot-parse';
import { INLINE_EQUATION_INLINE, MENTION_INLINE, SOURCE_INLINE } from './sciobot-inline';
import {
  BLOCK_EQUATION_MARKDOWN,
  CODE_BLOCK_MARKDOWN,
  INLINE_EQUATION_MARKDOWN,
} from './sciobot-markdown';
import {
  createMdxComponentRegistry,
  type MdxInlineBinding,
  type MdxMarkdownBinding,
  type MdxComponentDefinition,
  type MdxComponentRegistry,
} from './contract';

/** Payload types that sit inside a paragraph as one native atom. */
const INLINE_TYPES = new Set(['inline_equation', 'mention', 'inline_image', 'source']);
/** Payload types without MDX syntax of their own, kept as opaque application data. */
const OPAQUE_TYPES = ['listening'];
const MARKDOWN: Readonly<Record<string, MdxMarkdownBinding>> = {
  code_block: CODE_BLOCK_MARKDOWN,
  equation: BLOCK_EQUATION_MARKDOWN,
  inline_equation: INLINE_EQUATION_MARKDOWN,
};
const INLINE: Readonly<Record<string, MdxInlineBinding>> = {
  inline_equation: INLINE_EQUATION_INLINE,
  mention: MENTION_INLINE,
  source: SOURCE_INLINE,
};
const extras = (type: string) => ({
  ...(Object.hasOwn(INLINE, type) ? { inline: INLINE[type]! } : {}),
  ...(Object.hasOwn(SCIOBOT_DIFF_STRATEGIES, type) ? { diff: SCIOBOT_DIFF_STRATEGIES[type]! } : {}),
  ...(Object.hasOwn(SCIOBOT_PARSERS, type) ? { parse: SCIOBOT_PARSERS[type]! } : {}),
  ...(Object.hasOwn(MARKDOWN, type) ? { markdown: MARKDOWN[type]! } : {}),
});

function definitions(): MdxComponentDefinition[] {
  const out: MdxComponentDefinition[] = [];
  for (const [type, mdx] of DOMAIN_BLOCK_TAGS) {
    out.push({
      type,
      kind: INLINE_TYPES.has(type) ? 'inline' : 'block',
      content: mdx.kind === 'container' ? 'native' : 'none',
      mdx,
      ...extras(type),
    });
  }
  for (const type of OPAQUE_TYPES) {
    out.push({ type, kind: 'block', content: 'opaque-mdx', ...extras(type) });
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

let built: MdxComponentRegistry | undefined;

/**
 * The Sciobot components, storage names, parsers, and opaque MDX tags in one registry.
 *
 * It reads the existing adapter constants, so documents saved before the registry existed
 * keep their storage names. It is built on first use, so adapter modules can default to it
 * without an import cycle at load time.
 */
export function sciobotMdxRegistry(): MdxComponentRegistry {
  return (built ??= build());
}
