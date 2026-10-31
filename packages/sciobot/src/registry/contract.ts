import type { BlockTagDefinition } from '../projection/block-tags';
import type { DomainDiffStrategy } from '../diff/domain-strategies';
import type { TargetAstNode } from '../projection/parse-target';
import type { BlockPayload } from '../blocks/payload-schema';
import type { JsonObject, JsonValue } from '../data/json';

/** One MDX JSX attribute as the parser reads it. Expressions stay inert JSON objects. */
export interface MdxAttribute {
  readonly name: string;
  readonly value: JsonValue | undefined;
}

/** How a component reads its MDX element back into payload attributes. */
export interface MdxComponentParser {
  /** The JSX tag name this component reads. */
  readonly tag: string;
  /** The payload attributes of one element, or null when the element is not this component. */
  readonly read: (attributes: readonly MdxAttribute[], ast: TargetAstNode) => JsonObject | null;
}

/**
 * A component written as a Markdown node instead of a JSX element, such as a fenced code block
 * or a math block. Both hooks see the untrusted parsed node.
 */
export interface MdxMarkdownBinding {
  readonly types: readonly string[];
  /** The payload a new node describes, or null when it describes none. */
  readonly fromMarkdown: (ast: TargetAstNode, id: string) => BlockPayload | null;
  /** Whether a node, Markdown or JSX, still describes this payload exactly. */
  readonly matches: (ast: TargetAstNode, payload: BlockPayload) => boolean;
}

/**
 * A component that sits inside a paragraph as one native inline control.
 *
 * Both hooks may throw `MdxRefusal` to refuse with a stable message.
 */
export interface MdxInlineBinding {
  /** The atom name plans record for this control. */
  readonly atom: string;
  /** Markdown node types that write this atom inside a paragraph, such as `inlineMath`. */
  readonly types?: readonly string[];
  /** The payload of a new Markdown node of one of `types`. */
  readonly fromMarkdown?: (ast: TargetAstNode, id: string) => BlockPayload;
  /** The payload to store and the Word-readable text that stands for it. */
  readonly fallback: (payload: BlockPayload) => {
    readonly payload: BlockPayload;
    readonly text: string;
  };
}

/**
 * One application component that the embedding application registers.
 *
 * `type` is the payload type that a block stores. `mdx` maps the payload to its MDX tag and
 * back. A component without `mdx` has no MDX syntax of its own; its payload stays opaque.
 */
export interface MdxComponentDefinition {
  readonly type: string;
  /** A block owns paragraphs; an inline component is one native atom inside a paragraph. */
  readonly kind: 'block' | 'inline';
  /** `native`: the component wraps editable Word content. `none`: payload only. */
  readonly content: 'none' | 'native' | 'opaque-mdx';
  readonly mdx?: BlockTagDefinition;
  readonly parse?: MdxComponentParser;
  readonly markdown?: MdxMarkdownBinding;
  readonly inline?: MdxInlineBinding;
  readonly diff?: DomainDiffStrategy;
}

/** Where the adapter stores payloads in the DOCX package. One embedder owns one namespace. */
export interface MdxComponentStorage {
  /** The content-control tag prefix, such as `app:1:`. */
  readonly tagPrefix: string;
  /** The custom XML namespace URI of the payload part. */
  readonly namespace: string;
  /** The local name of the payload part root. */
  readonly root: string;
  /** The prefix of payload node IDs in the custom XML part. */
  readonly nodePrefix: string;
  /** The prefix of property review node IDs. */
  readonly reviewNodePrefix: string;
  /** The content-control tag of generated fallback label paragraphs. */
  readonly labelTag: string;
}

/** MDX tags the projection keeps as opaque source, by tag name and allowed attributes. */
export interface MdxOpaqueTags {
  readonly allowed: Readonly<Record<string, readonly string[]>>;
  /** Tags written back literally, without re-serializing their attributes. */
  readonly literal: readonly string[];
  /** Tags whose content is data, not document text. */
  readonly dataOnly: readonly string[];
}

export interface MdxRegistryInput {
  readonly storage: MdxComponentStorage;
  readonly components: readonly MdxComponentDefinition[];
  readonly opaque?: MdxOpaqueTags;
}

export interface MdxComponentRegistry {
  readonly storage: MdxComponentStorage;
  readonly components: ReadonlyMap<string, MdxComponentDefinition>;
  /** MDX mappings by payload type, the shape `projectBlocks` and `projectMdx` read. */
  readonly blockTags: ReadonlyMap<string, BlockTagDefinition>;
  readonly diffStrategies: ReadonlyMap<string, DomainDiffStrategy>;
  /**
   * Components by the JSX tag they read. One tag may name an atom and a container; the parser
   * picks the container when the element holds text.
   */
  readonly parsers: ReadonlyMap<
    string,
    { readonly atomic?: MdxComponentDefinition; readonly container?: MdxComponentDefinition }
  >;
  /** Components written as Markdown nodes, by node type. */
  readonly markdown: ReadonlyMap<string, MdxComponentDefinition>;
  /** Inline components written as Markdown nodes inside a paragraph, by node type. */
  readonly inlineMarkdown: ReadonlyMap<string, MdxComponentDefinition>;
  readonly opaque: {
    readonly allowed: ReadonlyMap<string, readonly string[]>;
    readonly literal: ReadonlySet<string>;
    readonly dataOnly: ReadonlySet<string>;
  };
}

export type MdxRegistryResult =
  | { readonly ok: true; readonly registry: MdxComponentRegistry }
  | { readonly ok: false; readonly reason: MdxRegistryRefusal; readonly detail: string };

export type MdxRegistryRefusal =
  | 'invalid-storage'
  | 'invalid-component'
  | 'duplicate-component'
  | 'duplicate-tag'
  | 'duplicate-parser'
  | 'invalid-opaque-tag';

const MAX_COMPONENTS = 512;
const MAX_NAME = 64;
const RESERVED = new Set(['__proto__', 'constructor', 'prototype']);
const TYPE = /^[a-z][a-z0-9_]*$/;
const TAG = /^[A-Za-z][A-Za-z0-9_-]*$/;
const ATTRIBUTE = /^[A-Za-z][A-Za-z0-9_-]*$/;
const TAG_PREFIX = /^[a-z][a-z0-9-]{0,15}:[0-9]{1,4}:$/;
const ID_PREFIX = /^[a-z][a-z0-9_-]{0,15}[_-]$/;
const LABEL = /^[a-z][a-z0-9-]{0,15}:[a-z][a-z0-9-]{0,31}$/;
const NAMESPACE = /^(urn:[a-z0-9][a-z0-9-]{0,31}(:[A-Za-z0-9._~-]+)+|https?:\/\/[^\s"'<>]+)$/;

const named = (value: unknown, pattern: RegExp): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= MAX_NAME &&
  !RESERVED.has(value) &&
  pattern.test(value);

function storageProblem(storage: MdxComponentStorage): string | null {
  if (!named(storage.tagPrefix, TAG_PREFIX)) return 'tagPrefix';
  if (
    typeof storage.namespace !== 'string' ||
    storage.namespace.length > 256 ||
    !NAMESPACE.test(storage.namespace)
  )
    return 'namespace';
  if (!named(storage.root, TYPE)) return 'root';
  if (!named(storage.nodePrefix, ID_PREFIX)) return 'nodePrefix';
  if (
    !named(storage.reviewNodePrefix, ID_PREFIX) ||
    storage.reviewNodePrefix === storage.nodePrefix
  )
    return 'reviewNodePrefix';
  if (!named(storage.labelTag, LABEL) || storage.labelTag.startsWith(storage.tagPrefix))
    return 'labelTag';
  return null;
}

function componentProblem(component: MdxComponentDefinition): string | null {
  if (!named(component.type, TYPE)) return 'type';
  if (component.kind !== 'block' && component.kind !== 'inline') return 'kind';
  if (!['none', 'native', 'opaque-mdx'].includes(component.content)) return 'content';
  if (component.content === 'native' && component.kind !== 'block') return 'content';
  const mdx = component.mdx;
  if (mdx) {
    if (!named(mdx.tag, TAG)) return 'mdx.tag';
    if (mdx.kind !== 'container' && mdx.kind !== 'atomic') return 'mdx.kind';
    if ((mdx.kind === 'container') !== (component.content === 'native')) return 'mdx.kind';
    if (typeof mdx.attributes !== 'function') return 'mdx.attributes';
  }
  if (component.parse) {
    if (!named(component.parse.tag, TAG)) return 'parse.tag';
    if (typeof component.parse.read !== 'function') return 'parse.read';
  }
  if (component.markdown) {
    const binding = component.markdown;
    if (!Array.isArray(binding.types) || binding.types.length === 0) return 'markdown.types';
    for (const type of binding.types) if (!named(type, TAG)) return 'markdown.types';
    if (typeof binding.fromMarkdown !== 'function' || typeof binding.matches !== 'function')
      return 'markdown.hooks';
  }
  if (component.inline) {
    const binding = component.inline;
    if (component.kind !== 'inline') return 'inline.kind';
    if (!named(binding.atom, TAG) || typeof binding.fallback !== 'function') return 'inline';
    for (const type of binding.types ?? []) if (!named(type, TAG)) return 'inline.types';
    if (binding.types?.length && typeof binding.fromMarkdown !== 'function')
      return 'inline.fromMarkdown';
  }
  const diff = component.diff;
  if (diff) {
    if (diff.kind !== 'atomic' && diff.kind !== 'container') return 'diff.kind';
    if (diff.kind === 'container')
      for (const key of [...(diff.identityProps ?? []), ...(diff.updateProps ?? [])])
        if (!named(key, ATTRIBUTE)) return 'diff.props';
  }
  return null;
}

/**
 * Build a frozen registry, or refuse with the first problem.
 *
 * Names are bounded and refuse prototype keys, because they become map keys, XML names, and
 * content-control tags. Two components may share an MDX tag only with different tag kinds,
 * as an atomic citation and a citation container do.
 */
export function createMdxComponentRegistry(input: MdxRegistryInput): MdxRegistryResult {
  const refuse = (reason: MdxRegistryRefusal, detail: string): MdxRegistryResult => ({
    ok: false,
    reason,
    detail,
  });
  const storageIssue = storageProblem(input.storage);
  if (storageIssue) return refuse('invalid-storage', storageIssue);
  if (input.components.length > MAX_COMPONENTS)
    return refuse('invalid-component', 'too many components');
  const components = new Map<string, MdxComponentDefinition>();
  const blockTags = new Map<string, BlockTagDefinition>();
  const diffStrategies = new Map<string, DomainDiffStrategy>();
  const tags = new Set<string>();
  const parsers = new Map<
    string,
    { atomic?: MdxComponentDefinition; container?: MdxComponentDefinition }
  >();
  const markdown = new Map<string, MdxComponentDefinition>();
  const inlineMarkdown = new Map<string, MdxComponentDefinition>();
  for (const component of input.components) {
    const issue = componentProblem(component);
    if (issue) return refuse('invalid-component', `${String(component.type)}: ${issue}`);
    if (components.has(component.type)) return refuse('duplicate-component', component.type);
    if (component.mdx) {
      const key = `${component.mdx.kind}:${component.mdx.tag}`;
      if (tags.has(key)) return refuse('duplicate-tag', key);
      tags.add(key);
      blockTags.set(component.type, Object.freeze({ ...component.mdx }));
    }
    if (component.diff) diffStrategies.set(component.type, Object.freeze({ ...component.diff }));
    const frozen = Object.freeze({ ...component });
    if (component.parse) {
      const slot = component.content === 'native' ? 'container' : 'atomic';
      const entry = parsers.get(component.parse.tag) ?? {};
      if (entry[slot]) return refuse('duplicate-parser', `${slot}:${component.parse.tag}`);
      parsers.set(component.parse.tag, { ...entry, [slot]: frozen });
    }
    for (const type of component.markdown?.types ?? []) {
      if (markdown.has(type)) return refuse('duplicate-parser', `markdown:${type}`);
      markdown.set(type, frozen);
    }
    for (const type of component.inline?.types ?? []) {
      if (inlineMarkdown.has(type)) return refuse('duplicate-parser', `inline:${type}`);
      inlineMarkdown.set(type, frozen);
    }
    components.set(component.type, frozen);
  }
  const opaque = input.opaque ?? { allowed: {}, literal: [], dataOnly: [] };
  const allowed = new Map<string, readonly string[]>();
  for (const [tag, attributes] of Object.entries(opaque.allowed)) {
    if (!named(tag, TAG)) return refuse('invalid-opaque-tag', tag);
    if (!Array.isArray(attributes) || attributes.some((name) => !named(name, ATTRIBUTE)))
      return refuse('invalid-opaque-tag', tag);
    allowed.set(tag, Object.freeze([...attributes]));
  }
  for (const tag of [...opaque.literal, ...opaque.dataOnly])
    if (!allowed.has(tag)) return refuse('invalid-opaque-tag', tag);
  return {
    ok: true,
    registry: Object.freeze({
      storage: Object.freeze({ ...input.storage }),
      components,
      blockTags,
      diffStrategies,
      parsers,
      markdown,
      inlineMarkdown,
      opaque: Object.freeze({
        allowed,
        literal: new Set(opaque.literal),
        dataOnly: new Set(opaque.dataOnly),
      }),
    }),
  };
}
