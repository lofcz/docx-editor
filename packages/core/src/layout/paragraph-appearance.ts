import {
  contentControlPropertiesOf,
  paragraphOffsetIndex,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlProperty,
} from '@docx-editor.dev/core/store';
import type {
  ParagraphAppearanceContribution,
  ParagraphAppearanceControl,
  ParagraphAppearanceDefaults,
} from '../contracts/paragraph-appearance.ts';
import type { StyleCascadeTable } from './style-cascade.ts';
import { stableHash } from '../store/comparators/canonical.ts';
import {
  revisionAttributionOf,
  revisionsVisible,
  withRevision,
  type RevisionAttribution,
  type RevisionAuthorFilter,
  type RevisionDisplayMode,
} from './revision-projection.ts';

const callbacks = new WeakMap<object, number>();
let nextCallback = 0;
interface AppearanceState {
  readonly contribution: ParagraphAppearanceContribution;
  readonly mode: RevisionDisplayMode;
  readonly filter: RevisionAuthorFilter | undefined;
  readonly paragraphs: WeakMap<OoxmlElement, readonly OoxmlProperty[] | null>;
}
const states = new WeakMap<StyleCascadeTable, AppearanceState>();
const wrapped = new WeakMap<StyleCascadeTable, Map<string, StyleCascadeTable>>();
const inherited = new WeakMap<
  readonly OoxmlProperty[],
  {
    readonly base: readonly OoxmlProperty[];
    readonly defaults: readonly OoxmlProperty[];
  }
>();
const themes = new Set([
  'dark1',
  'light1',
  'dark2',
  'light2',
  'text1',
  'background1',
  'text2',
  'background2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hyperlink',
  'followedHyperlink',
]);

/** Validate trusted host output before it enters the native Word property resolver. */
function propertiesOf(value: ParagraphAppearanceDefaults | null): readonly OoxmlProperty[] | null {
  if (value === null) return null;
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((key) => !['strike', 'color'].includes(key))
  )
    throw new TypeError('Paragraph appearance contains unsupported defaults.');
  const properties: OoxmlProperty[] = [];
  if (value.strike !== undefined) {
    if (typeof value.strike !== 'boolean')
      throw new TypeError('Paragraph appearance strike must be boolean.');
    properties.push({ localName: 'strike', attributes: { val: value.strike ? '1' : '0' } });
  }
  if (value.color !== undefined) {
    const color = value.color;
    if (
      !color ||
      typeof color !== 'object' ||
      Array.isArray(color) ||
      Object.keys(color).some((key) => !['val', 'themeColor', 'themeTint'].includes(key)) ||
      typeof color.val !== 'string' ||
      !/^(?:[\da-fA-F]{6}|auto)$/.test(color.val) ||
      (color.themeColor !== undefined && !themes.has(color.themeColor)) ||
      (color.themeTint !== undefined &&
        (typeof color.themeTint !== 'string' || !/^[\da-fA-F]{2}$/.test(color.themeTint)))
    )
      throw new TypeError('Paragraph appearance color must use bounded Word color values.');
    properties.push({ localName: 'color', attributes: { ...color } });
  }
  for (const property of properties) {
    Object.freeze(property.attributes);
    Object.freeze(property);
  }
  return Object.freeze(properties);
}

function controlsOf(
  paragraph: OoxmlElement,
  mode: RevisionDisplayMode,
  filter: RevisionAuthorFilter | undefined
): readonly ParagraphAppearanceControl[] {
  if (paragraph.kind !== 'paragraph') return Object.freeze([]);
  const offsets = paragraphOffsetIndex(paragraph);
  const controls: ParagraphAppearanceControl[] = [];
  const queue: { node: OoxmlNode; depth: number; revisions: readonly RevisionAttribution[] }[] = [
    { node: paragraph, depth: 0, revisions: [] },
  ];
  let count = 0;
  while (queue.length) {
    const entry = queue.pop()!;
    if (++count > 100_000 || entry.depth > 64)
      throw new RangeError('Paragraph appearance source exceeds its traversal limit.');
    const { node } = entry;
    if (node.kind === 'textValue') continue;
    const attribution = revisionAttributionOf(node);
    const revisions = attribution ? withRevision(entry.revisions, attribution) : entry.revisions;
    if (!revisionsVisible(revisions, mode === 'all-markup' ? 'proposed' : mode, filter)) continue;
    if (node.kind === 'contentControl') {
      const range = offsets.spanOf(node);
      if (range) {
        if (controls.length >= 256)
          throw new RangeError('Paragraph appearance has too many controls.');
        const properties = contentControlPropertiesOf(node);
        controls.push(
          Object.freeze({
            id: node.id,
            ...(properties.tag === undefined ? {} : { tag: properties.tag }),
            controlType: properties.type,
            start: range.start,
            end: range.end,
            ...(properties.checkbox ? { checked: properties.checkbox.checked } : {}),
          })
        );
      }
    }
    // A run cannot own an inline SDT. Skip its text and formatting descendants.
    if (node.kind === 'run' || node.kind === 'paragraphProperties') continue;
    for (let index = node.children.length - 1; index >= 0; index--)
      queue.push({ node: node.children[index]!, depth: entry.depth + 1, revisions });
  }
  return Object.freeze(controls);
}

/** A view-specific cascade keeps the existing paragraph/table/furniture pipeline intact. */
export function withParagraphAppearance(
  table: StyleCascadeTable,
  contribution: ParagraphAppearanceContribution | undefined,
  mode: RevisionDisplayMode,
  filter?: RevisionAuthorFilter
): StyleCascadeTable {
  if (!contribution) return table;
  if (
    typeof contribution.id !== 'string' ||
    !contribution.id.trim() ||
    contribution.id.length > 256 ||
    typeof contribution.resolve !== 'function'
  )
    throw new TypeError('Invalid paragraph appearance contribution.');
  let identity = callbacks.get(contribution.resolve);
  if (identity === undefined) {
    identity = ++nextCallback;
    callbacks.set(contribution.resolve, identity);
  }
  // Callback identity prevents a reused host ID from reusing a previous implementation's cache.
  const key = `${identity}:${mode}:${stableHash(filter?.cacheKey ?? '')}`;
  let entries = wrapped.get(table);
  if (!entries) {
    entries = new Map();
    wrapped.set(table, entries);
  }
  const existing = entries.get(key);
  if (existing) return existing;
  const result = { ...table, cacheToken: `${table.cacheToken}|appearance:${key}` };
  states.set(result, {
    contribution: Object.freeze({ id: contribution.id, resolve: contribution.resolve }),
    mode,
    filter,
    paragraphs: new WeakMap(),
  });
  if (entries.size >= 16) entries.delete(entries.keys().next().value!);
  entries.set(key, result);
  return result;
}

/** Attach defaults to the inherited properties without altering canonical nodes or style IDs. */
export function withParagraphAppearanceDefaults(
  table: StyleCascadeTable | undefined,
  paragraph: OoxmlElement,
  base: readonly OoxmlProperty[]
): readonly OoxmlProperty[] {
  const state = table && states.get(table);
  if (!state) return base;
  let defaults = state.paragraphs.get(paragraph);
  if (defaults === undefined) {
    defaults = propertiesOf(
      state.contribution.resolve(
        Object.freeze({
          paragraphId: paragraph.id,
          displayMode: state.mode,
          controls: controlsOf(paragraph, state.mode, state.filter),
        })
      )
    );
    state.paragraphs.set(paragraph, defaults);
  }
  if (!defaults?.length) return base;
  const result = [...base, ...defaults];
  inherited.set(result, { base, defaults });
  return result;
}

export const paragraphAppearanceOf = (properties: readonly OoxmlProperty[]) =>
  inherited.get(properties);

/** Explicit character-style and direct run properties keep their native precedence. */
export function applyParagraphAppearanceDefaults(
  resolved: readonly OoxmlProperty[],
  defaults: readonly OoxmlProperty[] | undefined,
  character: readonly OoxmlProperty[],
  direct: readonly OoxmlProperty[]
): readonly OoxmlProperty[] {
  if (!defaults?.length) return resolved;
  const authored = new Set([...character, ...direct].map((property) => property.localName));
  const additions = defaults.filter((property) => !authored.has(property.localName));
  return additions.length ? [...resolved, ...additions] : resolved;
}
