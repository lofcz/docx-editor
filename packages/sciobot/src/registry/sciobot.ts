import { DOMAIN_BLOCK_TAGS } from '../projection/block-tags';
import { SCIOBOT_DIFF_STRATEGIES } from './sciobot-diff';
import {
  CUSTOM_ALLOWED_TAGS,
  CUSTOM_DATA_ONLY_TAGS,
  CUSTOM_LITERAL_TAGS,
} from '../projection/mdx-tags';
import { SCIOBOT_STORAGE } from './storage';
import { SCIOBOT_PARSERS } from './sciobot-parse';
import { INLINE_IMAGE } from '../blocks/inline-image-codec';
import { SCIOBOT_VALIDATORS } from './sciobot-valid';
import { SCIOBOT_WORD_FALLBACKS } from './sciobot-fallback';
import {
  DATE_INLINE,
  INLINE_EQUATION_INLINE,
  MENTION_INLINE,
  SOURCE_INLINE,
} from './sciobot-inline';
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
const INLINE_TYPES = new Set(['inline_equation', 'mention', 'inline_image', 'source', 'date']);
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
  date: DATE_INLINE,
};
/** An inline image owns the media part its payload names. */
const imageMedia = (payload: { readonly attributes: Readonly<Record<string, unknown>> }) =>
  typeof payload.attributes.mediaPartName === 'string' ? payload.attributes.mediaPartName : null;
/** Recap sections list their entries in the application's own style. */
const RECAP_RULES: Readonly<Record<string, MdxComponentDefinition['contentRules']>> = {
  recap_goals: { suppressLists: true, stripNumberPrefix: true },
  recap_curriculum: { suppressLists: true },
  recap_materials: { suppressLists: true },
  recap_adaptations: { suppressLists: true },
};
/** Citation serializers drop a book reference or trust level they cannot read. */
const SOURCE_PROBE = { bookId: 1, trust: 'verified' } as const;
const extras = (type: string) => ({
  ...(type === 'source' || type === 'source_excerpt' ? { propertyProbe: SOURCE_PROBE } : {}),
  ...(Object.hasOwn(RECAP_RULES, type) ? { contentRules: RECAP_RULES[type]! } : {}),
  ...(Object.hasOwn(SCIOBOT_WORD_FALLBACKS, type)
    ? { wordFallback: SCIOBOT_WORD_FALLBACKS[type]! }
    : {}),
  ...(type === INLINE_IMAGE ? { mediaPart: imageMedia } : {}),
  ...(Object.hasOwn(SCIOBOT_VALIDATORS, type) ? { valid: SCIOBOT_VALIDATORS[type]! } : {}),
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
    storage: SCIOBOT_STORAGE,
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
