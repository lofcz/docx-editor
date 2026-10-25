import { serializeInlineImage } from './inline-image';
import { serializeMath } from './math';
import { serializeMention } from './mention';
import { serializeCodeBlock } from './code-block';
import type { BlockPayload, JsonValue } from '../blocks/payload-schema';

export interface BlockAttributeBinding {
  readonly name: string;
  readonly property: string;
  readonly value: JsonValue;
}

/** A path addresses persisted JSON without interpreting names as JavaScript properties. */
export type BlockPropertyPath = readonly (string | number)[];

export interface SerializedBlockProperty {
  readonly name: string;
  readonly property: string;
  readonly path: BlockPropertyPath;
  readonly start: number;
  readonly end: number;
  readonly valueStart: number | null;
  readonly valueEnd: number | null;
}

export interface SerializedBlockTag {
  readonly tag: string;
  readonly text: string;
  readonly properties: readonly SerializedBlockProperty[];
}

export interface BlockTagDefinition {
  readonly tag: string;
  readonly kind: 'container' | 'atomic';
  readonly attributes: (payload: BlockPayload) => readonly BlockAttributeBinding[];
  /** Atomic payloads can serialize complete stored syntax, including opaque MDX trees. */
  readonly serialize?: (payload: BlockPayload) => SerializedBlockTag;
  /** Rich serializers default to HTML cells. Literal inline syntax can explicitly retain GFM. */
  readonly serializedCellSyntax?: 'gfm' | 'html';
}

function all(payload: BlockPayload): BlockAttributeBinding[] {
  return Object.entries(payload.attributes)
    .filter(
      ([key]) =>
        !['id', 'type', 'children', 'suggestion', 'suggestionTransient', 'listStart'].includes(
          key
        ) &&
        !key.startsWith('_') &&
        !key.startsWith('suggestion_')
    )
    .map(([property, value]) => ({
      name: property,
      property,
      value,
    }));
}
function renamed(
  payload: BlockPayload,
  names: Readonly<Record<string, string>>
): BlockAttributeBinding[] {
  return all(payload).map((binding) => ({
    ...binding,
    name: Object.hasOwn(names, binding.property) ? names[binding.property]! : binding.name,
  }));
}
const empty = () => [];
function ordered(payload: BlockPayload, keys: readonly string[]): BlockAttributeBinding[] {
  return keys
    .filter((key) => Object.hasOwn(payload.attributes, key))
    .map((property) => ({ name: property, property, value: payload.attributes[property]! }));
}
function title(payload: BlockPayload): BlockAttributeBinding[] {
  const keys = ['grade', 'organizationalMinutes', 'lessonMinutes'];
  return [
    ...ordered(payload, keys)
      .filter((binding) => binding.property !== 'grade' || Boolean(binding.value))
      .map((binding) => ({
        ...binding,
        name:
          binding.property === 'organizationalMinutes'
            ? 'organizational_minutes'
            : binding.property === 'lessonMinutes'
              ? 'lesson_minutes'
              : binding.name,
      })),
    ...all(payload).filter((binding) => !keys.includes(binding.property)),
  ];
}
function visual(payload: BlockPayload): BlockAttributeBinding[] {
  const keys = ['kind', 'phase', 'src', 'alt', 'prompt'];
  return [
    ...ordered(payload, keys).filter(
      (binding) =>
        binding.property !== 'phase' ||
        (payload.attributes.kind === 'explainer' && Boolean(binding.value))
    ),
    ...all(payload).filter((binding) => !keys.includes(binding.property)),
  ];
}
const phase = (payload: BlockPayload) => [
  ...renamed(payload, { phaseType: 'type' }).filter((binding) => binding.property === 'phaseType'),
  ...all(payload).filter((binding) => binding.property !== 'phaseType'),
];
function source(payload: BlockPayload): BlockAttributeBinding[] {
  const out: BlockAttributeBinding[] = [];
  const attrs = payload.attributes;
  for (const [property, name] of [
    ['trust', 'type'],
    ['url', 'url'],
    ['title', 'title'],
    ['author', 'author'],
    ['bookId', 'bookId'],
  ] as const) {
    if (!Object.hasOwn(attrs, property)) continue;
    const value = attrs[property]!;
    if (property === 'trust' && value !== 'verified' && value !== 'generated') continue;
    if (
      property === 'bookId'
        ? typeof value !== 'number' || value <= 0
        : typeof value !== 'string' || value.length === 0
    )
      continue;
    out.push({ property, name, value: property === 'bookId' ? String(value) : value });
  }
  const known = new Set(['trust', 'url', 'title', 'author', 'bookId', 'bookid', 'sourceType']);
  for (const binding of all(payload)) {
    if (
      !known.has(binding.property) &&
      typeof binding.value === 'string' &&
      binding.value.length > 0
    )
      out.push(binding);
    if (
      !known.has(binding.property) &&
      typeof binding.value === 'number' &&
      binding.value > 0 &&
      binding.property.toLowerCase() === 'bookid'
    ) {
      const value = { ...binding, name: 'bookId', value: String(Math.trunc(binding.value)) };
      const previous = out.findIndex((entry) => entry.name === 'bookId');
      if (previous < 0) out.push(value);
      else out[previous] = value;
    }
  }
  return out;
}

/** Application container rules serialize remaining attributes in insertion order. */
function containerAttributes(payload: BlockPayload): BlockAttributeBinding[] {
  return Object.entries(payload.attributes)
    .filter(([property]) => !['id', 'type', 'children'].includes(property))
    .map(([property, value]) => ({ name: property, property, value }));
}
function callout(payload: BlockPayload): BlockAttributeBinding[] {
  const attributes = Object.entries(payload.attributes)
    .filter(([property]) => !['type', 'children', 'variant'].includes(property))
    .map(([property, value]) => ({ name: property, property, value }));
  const variant = payload.attributes.variant;
  if (typeof variant === 'string' && variant.length > 0) {
    attributes.push({ name: 'type', property: 'variant', value: variant });
  }
  return attributes;
}

/** Same tag names and property remapping as the application Markdown rules. */
export const DOMAIN_BLOCK_TAGS: ReadonlyMap<string, BlockTagDefinition> = new Map<
  string,
  BlockTagDefinition
>([
  [
    'code_block',
    {
      tag: 'code',
      kind: 'atomic',
      attributes: (payload) => ordered(payload, ['language', 'value']),
      serialize: serializeCodeBlock,
    },
  ],
  ['callout', { tag: 'callout', kind: 'container', attributes: callout }],
  ...['story', 'solution'].map(
    (tag) => [tag, { tag, kind: 'container', attributes: containerAttributes }] as const
  ),
  [
    'lesson_title',
    {
      tag: 'lesson_title',
      kind: 'container',
      attributes: title,
    },
  ],
  ...['preparation_recap', 'recap_goals', 'recap_materials', 'recap_adaptations'].map(
    (tag) => [tag, { tag, kind: 'container', attributes: empty }] as const
  ),
  [
    'recap_curriculum',
    {
      tag: 'recap_curriculum',
      kind: 'container',
      attributes: (payload: BlockPayload) =>
        all(payload).filter(
          (binding) => binding.property !== 'framework' || Boolean(binding.value)
        ),
    },
  ],
  [
    'recap_phase',
    {
      tag: 'recap_phase',
      kind: 'container',
      attributes: (payload: BlockPayload) =>
        phase(payload).filter((binding) => binding.property !== 'duration'),
    },
  ],
  ['lesson_phase', { tag: 'phase', kind: 'container', attributes: phase }],
  [
    'lesson_activity',
    {
      tag: 'activity',
      kind: 'container',
      attributes: (payload: BlockPayload) =>
        ordered(payload, ['name', 'duration', 'resource']).filter(
          (binding) => binding.property !== 'resource' || Boolean(binding.value)
        ),
    },
  ],
  ['lesson_visual', { tag: 'lesson_visual', kind: 'atomic', attributes: visual }],
  ['source', { tag: 'source', kind: 'atomic', attributes: source }],
  [
    'inline_equation',
    {
      tag: 'scio-inline-equation',
      kind: 'atomic',
      attributes: (payload) => ordered(payload, ['texExpression']),
      serialize: serializeMath,
      serializedCellSyntax: 'gfm',
    },
  ],
  [
    'equation',
    {
      tag: 'scio-equation',
      kind: 'atomic',
      attributes: (payload) => ordered(payload, ['texExpression']),
      serialize: serializeMath,
    },
  ],
  [
    'mention',
    {
      tag: 'mention',
      kind: 'atomic',
      attributes: (payload) => ordered(payload, ['key', 'value']),
      serialize: serializeMention,
      serializedCellSyntax: 'gfm',
    },
  ],
  [
    'inline_image',
    {
      tag: 'scio-inline-image',
      kind: 'atomic',
      attributes: empty,
      serialize: serializeInlineImage,
      serializedCellSyntax: 'gfm',
    },
  ],
  ['source_excerpt', { tag: 'source', kind: 'container', attributes: source }],
]);

/** A nested block SDT with this tag contains generated fallback label paragraphs. */
export const GENERATED_LABEL_TAG = 'scio:label';

/** Serialize data as quoted attributes, matching propsToAttributes' scalar/JSON behavior. */
export function attributeLiteral(value: JsonValue): string | null {
  if (value === null) return null;
  if (value === true) return '';
  const text = typeof value === 'string' ? value : JSON.stringify(value);
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('"', '&quot;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('\n', '&#10;')
    .replaceAll('\r', '&#13;');
}
