import type { AutomationHandleTable } from './handles.ts';
import { projectParagraphText, type ProjectedParagraphText } from './text-projection.ts';
import type { AutomationTextProjection } from './operations.ts';
// Word paragraph marks can be resolved while collaboration retains their physical fragments.
// Keep every editable offset in the original canonical fragment.
import type { OoxmlParagraphNode, OoxmlNode } from '../store/package/ooxml-tree.ts';
import { isJoinedParagraph } from '../store/package/retired-paragraph.ts';
import {
  paragraphMergeSources,
  tableParagraphMergeTarget,
} from '../store/store/paragraph-mark-merge.ts';
import type { AutomationStoryReads } from './reads.ts';
import type { ResolvedPoint, ResolvedRange } from './spans.ts';

interface Group {
  readonly id: string;
  readonly fragments: readonly string[];
}
interface Groups {
  readonly ordered: readonly Group[];
  readonly byFragment: ReadonlyMap<string, Group>;
}
const cache = new WeakMap<AutomationStoryReads, Groups>();

export function joinedParagraphGroups(reads: AutomationStoryReads): Groups {
  const cached = cache.get(reads);
  if (cached) return cached;
  const links = new Map<string, string>();
  const walk = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    const sources = paragraphMergeSources(node.children);
    let next: string | null = null;
    for (let index = node.children.length - 1; index >= 0; index--) {
      const child = node.children[index]!;
      if (child.kind === 'paragraph') {
        if (next && sources.has(child.id) && isJoinedParagraph(child)) links.set(child.id, next);
        next = child.id;
      } else if (child.kind === 'table') next = tableParagraphMergeTarget(child)?.id ?? null;
      else if (
        child.kind === 'contentControl' ||
        (child.kind === 'generic' &&
          child.namespaceUri === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' &&
          child.localName === 'sdt')
      )
        next = null;
    }
    for (const child of node.children) walk(child);
  };
  walk(reads.root);
  const grouped = new Map<string, string[]>();
  for (const id of reads.paragraphIds) {
    let owner = id;
    const seen = new Set<string>();
    while (links.has(owner) && !seen.has(owner)) {
      seen.add(owner);
      const target = links.get(owner)!;
      if (!reads.has(target)) break;
      owner = target;
    }
    const members = grouped.get(owner) ?? [];
    members.push(id);
    grouped.set(owner, members);
  }
  const ordered = [...grouped].map(([id, fragments]) => ({ id, fragments }));
  const byFragment = new Map<string, Group>();
  for (const group of ordered) for (const id of group.fragments) byFragment.set(id, group);
  const result = { ordered, byFragment };
  cache.set(reads, result);
  return result;
}

export function logicalParagraphIds(
  reads: AutomationStoryReads,
  ids = reads.paragraphIds
): string[] {
  const groups = joinedParagraphGroups(reads);
  return [...new Set(ids.map((id) => groups.byFragment.get(id)?.id ?? id))];
}

export function logicalParagraphRange(
  reads: AutomationStoryReads,
  id: string
): ResolvedRange | null {
  const group = joinedParagraphGroups(reads).byFragment.get(id);
  if (!group) return null;
  const point = (paragraphId: string, end: boolean): ResolvedPoint => ({
    story: reads.story,
    paragraphId,
    index: reads.indexOf(paragraphId),
    offset: end ? (reads.rawText(paragraphId) ?? '').length : 0,
  });
  return { start: point(group.fragments[0]!, false), end: point(group.fragments.at(-1)!, true) };
}

/** A logical offset at an old mark lands before the next fragment's first character. */
export function logicalParagraphPoint(
  reads: AutomationStoryReads,
  id: string,
  offset: number,
  edge: 'start' | 'end' = 'start'
): ResolvedPoint | null {
  if (!Number.isSafeInteger(offset) || offset < 0) return null;
  const group = joinedParagraphGroups(reads).byFragment.get(id);
  if (!group) return null;
  let remaining = offset;
  for (const [index, paragraphId] of group.fragments.entries()) {
    const length = (reads.rawText(paragraphId) ?? '').length;
    if (
      remaining < length ||
      (edge === 'end' && remaining === length) ||
      (index === group.fragments.length - 1 && remaining === length)
    )
      return {
        story: reads.story,
        paragraphId,
        index: reads.indexOf(paragraphId),
        offset: remaining,
      };
    remaining -= length;
  }
  return null;
}

export function hasJoinedBoundary(
  reads: AutomationStoryReads,
  left: string,
  right: string
): boolean {
  const groups = joinedParagraphGroups(reads);
  const group = groups.byFragment.get(left);
  if (!group || group !== groups.byFragment.get(right)) return false;
  return group.fragments.indexOf(right) === group.fragments.indexOf(left) + 1;
}

/** Offset in the joined Word paragraph, while the supplied point keeps its physical address. */
export function logicalParagraphOffset(
  reads: AutomationStoryReads,
  paragraphId: string,
  offset: number
): number {
  const group = joinedParagraphGroups(reads).byFragment.get(paragraphId);
  let before = 0;
  for (const id of group?.fragments ?? []) {
    if (id === paragraphId) return before + offset;
    before += (reads.rawText(id) ?? '').length;
  }
  return offset;
}

const projections = new WeakMap<AutomationStoryReads, Map<string, ProjectedParagraphText>>();
/** A read-only view uses the surviving mark's properties and every original content node. */
export function logicalParagraphProjection(
  reads: AutomationStoryReads,
  id: string,
  projection: AutomationTextProjection
): ProjectedParagraphText | null {
  const group = joinedParagraphGroups(reads).byFragment.get(id);
  if (!group) return null;
  if (group.fragments.length === 1) return reads.projectedText(group.id, projection);
  const key = `${group.id}:${projection}`;
  let cached = projections.get(reads);
  if (!cached) {
    cached = new Map();
    projections.set(reads, cached);
  }
  const found = cached.get(key);
  if (found) return found;
  const owner = reads.node(group.id);
  if (owner?.kind !== 'paragraph') return null;
  const properties = (node: OoxmlNode) => node.kind === 'paragraphProperties';
  const children = [
    ...owner.children.filter(properties),
    ...group.fragments.flatMap((fragment) => {
      const node = reads.node(fragment);
      return node?.kind === 'paragraph' ? node.children.filter((child) => !properties(child)) : [];
    }),
  ];
  const node = { ...owner, children } as OoxmlParagraphNode;
  const raw = group.fragments.map((fragment) => reads.rawText(fragment) ?? '').join('');
  const result = projectParagraphText(node, raw, projection);
  cached.set(key, result);
  return result;
}

export function isLogicalParagraph(reads: AutomationStoryReads, id: string): boolean {
  return (joinedParagraphGroups(reads).byFragment.get(id)?.fragments.length ?? 0) > 1;
}

export function logicalParagraphHandles(
  reads: AutomationStoryReads,
  handles: AutomationHandleTable,
  ids = reads.paragraphIds
) {
  return logicalParagraphIds(reads, ids).map((id) =>
    handles.paragraph(id, reads.story, isLogicalParagraph(reads, id))
  );
}

export function logicalParagraphSlice(
  reads: AutomationStoryReads,
  id: string,
  from: number,
  to: number
): ResolvedRange | null {
  const start = logicalParagraphPoint(reads, id, from),
    end = logicalParagraphPoint(reads, id, to, 'end');
  return start && end ? { start, end } : null;
}
