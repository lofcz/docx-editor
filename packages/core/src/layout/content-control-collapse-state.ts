import {
  findNode,
  parentNodeOf,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type {
  CollapsedContentControl,
  ContentControlCollapseResult,
} from '../contracts/content-control-collapse.ts';
import {
  collapseOwnerRefusal,
  prepareContentControlCollapse,
  type PreparedContentControlCollapse,
} from './content-control-collapse.ts';
import type { SemanticLayoutOptions } from './semantic-layout-options.ts';
import { contentControlPlaceholder } from './content-control-placeholder.ts';

/** One immutable input for visible body flow and its cache identity. @internal */
export type ContentControlCollapseLayout = NonNullable<
  SemanticLayoutOptions['contentControlCollapse']
>;

const flows = new WeakMap<
  ContentControlCollapseLayout,
  WeakMap<readonly OoxmlElement[], readonly OoxmlElement[]>
>();
const tableViews = new WeakMap<ContentControlCollapseLayout, WeakMap<OoxmlNode, OoxmlNode>>();

/** Filter table flow without changing canonical cells, properties, paragraph identities, or counters. */
function visibleTableNode(
  node: OoxmlNode,
  collapse: ContentControlCollapseLayout,
  cache: WeakMap<OoxmlNode, OoxmlNode>,
  depth = 0
): OoxmlNode {
  const known = cache.get(node);
  if (known) return known;
  // Validation admits owners only within this depth. Paragraph runs never contain block folds.
  if (
    depth > 64 ||
    !['table', 'tableRow', 'tableCell', 'contentControl', 'contentControlContent'].includes(
      node.kind
    )
  )
    return node;
  if (node.kind === 'textValue') return node;
  let changed = false;
  const children: OoxmlNode[] = [];
  for (const child of node.children) {
    const placeholder = collapse.placeholders?.get(child.id);
    if (placeholder) {
      changed = true;
      children.push(placeholder);
      continue;
    }
    if (collapse.hiddenIds.has(child.id) && !collapse.placeholderAncestors?.has(child.id)) {
      changed = true;
      continue;
    }
    const shown = visibleTableNode(child, collapse, cache, depth + 1);
    changed ||= shown !== child;
    children.push(shown);
  }
  const shown = changed ? ({ ...node, children } as OoxmlNode) : node;
  cache.set(node, shown);
  return shown;
}

/** Keep unchanged flow inputs stable while canonical list and field resolution stays separate. */
export function visibleCollapsedControlFlow(
  blocks: readonly OoxmlElement[],
  collapse?: ContentControlCollapseLayout
): readonly OoxmlElement[] {
  if (!collapse) return blocks;
  let cache = flows.get(collapse);
  if (!cache) flows.set(collapse, (cache = new WeakMap()));
  const cached = cache.get(blocks);
  if (cached) return cached;
  let tableCache = tableViews.get(collapse);
  if (!tableCache) tableViews.set(collapse, (tableCache = new WeakMap()));
  let changed = false;
  const visible: OoxmlElement[] = [];
  for (const block of blocks) {
    const placeholder = collapse.placeholders?.get(block.id);
    if (placeholder) {
      changed = true;
      visible.push(placeholder);
      continue;
    }
    if (collapse.hiddenIds.has(block.id)) {
      changed = true;
      continue;
    }
    const shown = visibleTableNode(block, collapse, tableCache) as OoxmlElement;
    changed ||= shown !== block;
    visible.push(shown);
  }
  const result = changed ? visible : blocks;
  cache.set(blocks, result);
  return result;
}

/** One surface owns this state. It contains canonical IDs and no editable text. @internal */
export function createContentControlCollapseState() {
  let entries: readonly PreparedContentControlCollapse[] = [];
  let snapshot: readonly CollapsedContentControl[] = Object.freeze([]);
  let layout: ContentControlCollapseLayout | undefined;
  let sourcePart: OoxmlPart | undefined;
  let epoch = 0;
  const adopt = (next: readonly PreparedContentControlCollapse[], part: OoxmlPart) => {
    const same =
      entries.length === next.length &&
      entries.every((entry, index) => {
        const other = next[index]!;
        return (
          entry.controlId === other.controlId &&
          entry.summaryParagraphId === other.summaryParagraphId
        );
      });
    entries = next;
    sourcePart = part;
    if (!same)
      snapshot = Object.freeze(
        next.map(({ controlId, summaryParagraphId }) =>
          Object.freeze({ controlId, summaryParagraphId })
        )
      );
    const hiddenIds = new Set<string>();
    for (const entry of next) for (const id of entry.hiddenIds) hiddenIds.add(id);
    const placeholders = new Map<string, OoxmlElement>();
    const placeholderAncestors = new Set<string>();
    for (const entry of next) {
      if (!entry.placeholder || hiddenIds.has(entry.controlId)) continue;
      placeholders.set(entry.placeholder.id, contentControlPlaceholder(entry.placeholder));
      let ancestor = parentNodeOf(part, entry.placeholder.id);
      while (ancestor && ancestor.id !== entry.controlId) {
        placeholderAncestors.add(ancestor.id);
        ancestor = parentNodeOf(part, ancestor.id);
      }
    }
    const unchanged =
      hiddenIds.size === (layout?.hiddenIds.size ?? 0) &&
      [...hiddenIds].every((id) => layout?.hiddenIds.has(id)) &&
      placeholders.size === (layout?.placeholders?.size ?? 0) &&
      [...placeholders].every(([id, node]) => layout?.placeholders?.get(id) === node) &&
      placeholderAncestors.size === (layout?.placeholderAncestors?.size ?? 0) &&
      [...placeholderAncestors].every((id) => layout?.placeholderAncestors?.has(id));
    if (!unchanged)
      layout = hiddenIds.size
        ? { token: `collapse:${++epoch}`, hiddenIds, placeholders, placeholderAncestors }
        : undefined;
    return !same || !unchanged;
  };
  return {
    current: () => snapshot,
    option: () => layout,
    set(
      part: OoxmlPart,
      requests: ReadonlyMap<string, string | null>
    ): ContentControlCollapseResult & { readonly changed?: boolean } {
      const prepared = prepareContentControlCollapse(part, requests);
      if (!prepared.ok) return prepared;
      return { ok: true, changed: adopt(prepared.entries, part) };
    },
    /** Prune invalid owners. Only explicit collaboration aliases can replace missing IDs. */
    reconcile(part: OoxmlPart, resolveNodeId?: (id: string) => string | undefined): boolean {
      if (sourcePart === part || !entries.length) {
        sourcePart = part;
        return false;
      }
      const next: PreparedContentControlCollapse[] = [];
      const budget = { visited: 0 };
      for (const entry of entries) {
        const controlId = findNode(part, entry.controlId)
          ? entry.controlId
          : resolveNodeId?.(entry.controlId);
        const summaryId =
          entry.summaryParagraphId === null
            ? null
            : findNode(part, entry.summaryParagraphId)
              ? entry.summaryParagraphId
              : resolveNodeId?.(entry.summaryParagraphId);
        if (!controlId || summaryId === undefined) continue;
        const owner = findNode(part, controlId)!;
        if (collapseOwnerRefusal(part, owner)) continue;
        if (sourcePart && findNode(sourcePart, entry.controlId) === owner) {
          next.push(entry);
          continue;
        }
        const prepared = prepareContentControlCollapse(
          part,
          new Map([[controlId, summaryId]]),
          budget
        );
        if (prepared.ok) next.push(prepared.entries[0]!);
      }
      return adopt(next, part);
    },
    /** Explicit local navigation reveals enclosing folds before caret placement. */
    reveal(part: OoxmlPart, paragraphIds: readonly string[]): boolean {
      const requested = paragraphIds.length > 2 ? new Set(paragraphIds) : null;
      const next = entries.filter((entry) => {
        if (!requested) return !paragraphIds.some((id) => entry.hiddenIds.has(id));
        for (const id of entry.hiddenIds) if (requested.has(id)) return false;
        return true;
      });
      return next.length !== entries.length && adopt(next, part);
    },
  };
}
