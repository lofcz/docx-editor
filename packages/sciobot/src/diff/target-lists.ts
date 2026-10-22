import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { LogicalBlock } from './logical-block-types';
import { isListTarget, type ListSourceState, type ListTarget } from './list-types';
import { listInstanceStart, planListMembership } from './lists';

export interface TargetListPair {
  readonly before: LogicalBlock;
  readonly afterIndex: number;
  readonly paragraphId: string;
}
export type TargetListResolution =
  | { readonly success: true; readonly targets: readonly ListTarget[] }
  | NativeEditFailure;
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
const groupKey = (block: LogicalBlock) => block.list?.group.join('/') ?? '';

/** Resolve syntax groups to existing/new native instances. Retained displayed ordinals never request restarts. */
export function resolveTargetLists(
  source: {
    readonly revision: RevisionToken;
    readonly pkg: OoxmlPackage;
    readonly partName: string;
  },
  after: readonly LogicalBlock[],
  pairs: readonly TargetListPair[]
): TargetListResolution {
  type Candidate = { index: number; paragraphId: string; numId: string };
  const candidates = new Map<string, Candidate[]>(),
    owners = new Map<number, Candidate>(),
    eligible: TargetListPair[] = [];
  for (const pair of pairs) {
    const old = pair.before.list,
      next = after[pair.afterIndex]?.list;
    if (
      pair.before.kind !== 'paragraph' ||
      after[pair.afterIndex]?.kind !== 'paragraph' ||
      !old ||
      !next ||
      old.checked !== null ||
      next.checked !== null ||
      old.suppressed ||
      next.suppressed ||
      old.role !== 'item' ||
      next.role !== 'item' ||
      old.marker !== next.marker
    )
      continue;
    eligible.push(pair);
  }
  // Read all authored/inherited membership through one D06 probe. Never apply or return it.
  const probe = planListMembership(
    source,
    eligible.map((pair) => ({ paragraphId: pair.paragraphId, target: { kind: 'none' } }))
  );
  if (!probe.success) return probe;
  const states = new Map<string, ListSourceState>();
  for (const edit of probe.plan.operations)
    if (edit.kind === 'list') states.set(edit.paragraphId, edit.before);
  for (const pair of eligible) {
    const state = states.get(pair.paragraphId);
    if (!state?.numbering)
      return fail('A retained source list has no canonical numbering instance.');
    const key = groupKey(after[pair.afterIndex]!),
      list = candidates.get(key) ?? [];
    const candidate = {
      index: pair.afterIndex,
      paragraphId: pair.paragraphId,
      numId: state.numbering.numId,
    };
    list.push(candidate);
    owners.set(pair.afterIndex, candidate);
    candidates.set(key, list);
  }
  for (const list of candidates.values()) list.sort((a, b) => a.index - b.index);
  const targets: ListTarget[] = [],
    groups = new Map<string, { target: ListTarget; anchor?: Candidate }>(),
    checked = new Set<string>(),
    starts = new Map<string, number | null>();
  for (let index = 0; index < after.length; index++) {
    const block = after[index]!,
      list = block.list;
    if (list && block.kind !== 'paragraph')
      return fail('Native list intentions require a paragraph owner.');
    let target: ListTarget;
    if (!list || (list.suppressed && list.checked === null)) target = { kind: 'none' };
    else if (list.role === 'continuation')
      return fail('List continuation paragraphs need explicit native indentation intentions.');
    else if (list.checked !== null)
      target = {
        kind: 'checklist',
        checked: list.checked,
        level: list.suppressed ? 0 : list.level,
      };
    else {
      const key = groupKey(block),
        anchors = candidates.get(key) ?? [];
      let anchor = owners.get(index);
      if (!anchor) {
        let low = 0,
          high = anchors.length;
        while (low < high) {
          const middle = (low + high) >>> 1;
          if (anchors[middle]!.index < index) low = middle + 1;
          else high = middle;
        }
        anchor = anchors[low - 1] ?? anchors[low];
      }
      let parent: { target: ListTarget; anchor?: Candidate } | undefined;
      for (let end = list.group.length - 1; end > 0 && !parent; end--)
        parent = groups.get(list.group.slice(0, end).join('/'));
      const ancestor = parent?.target;
      const inheritedAnchor = !anchor;
      if (!anchor && ancestor && (ancestor.kind === 'ordered' || ancestor.kind === 'bullet'))
        anchor = parent?.anchor;
      if (!anchor && list.level > 0 && !ancestor)
        return fail('Nested target numbering has no parent instance.');
      if (!anchor && ancestor?.kind === 'checklist')
        return fail('Numbered children of native checklists need explicit mixed-list ownership.');
      const inherited =
        ancestor && (ancestor.kind === 'ordered' || ancestor.kind === 'bullet')
          ? ancestor.instance
          : undefined;
      target = {
        kind: list.marker,
        level: list.level,
        instance: anchor
          ? { kind: 'existing', numId: anchor.numId }
          : {
              kind: 'new',
              key: inherited?.kind === 'new' ? inherited.key : `target-list:${key}`,
              ...(list.marker === 'ordered' ? { start: list.start ?? 1 } : {}),
            },
      };
      if (anchor) {
        if (inheritedAnchor && list.marker === 'ordered') {
          const key = JSON.stringify([anchor.numId, list.level]);
          if (!starts.has(key))
            starts.set(key, listInstanceStart(source.pkg, anchor.numId, list.level));
          if (starts.get(key) !== (list.start ?? 1))
            return fail('A new nested start would rewrite an existing Word numbering definition.');
        }
        const key = JSON.stringify([anchor.numId, target.kind, target.level]);
        if (!checked.has(key)) {
          const check = planListMembership(source, [{ paragraphId: anchor.paragraphId, target }]);
          if (!check.success) return check;
          checked.add(key);
        }
      }
      groups.set(key, { target, ...(anchor ? { anchor } : {}) });
    }
    if (list && target.kind === 'checklist') groups.set(groupKey(block), { target });
    if (!isListTarget(target)) return fail('The target list exceeds native level or start limits.');
    targets.push(target);
  }
  return { success: true, targets: Object.freeze(targets.map((target) => Object.freeze(target))) };
}
