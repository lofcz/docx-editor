import { diffArrays } from 'diff';
import type { NativeEditFailure } from '../contracts/document';
import { SIBLING_BLOCK_LIMITS } from './sibling-blocks';

export interface TargetSequenceAlignment {
  readonly pairs: readonly { readonly before: number; readonly after: number }[];
  readonly gaps: readonly {
    readonly beforeStart: number;
    readonly beforeEnd: number;
    readonly afterStart: number;
    readonly afterEnd: number;
  }[];
}
export const targetKeyCounts = (keys: readonly string[]) => {
  const counts = new Map<string, number>();
  for (const key of keys) counts.set(key, (counts.get(key) ?? 0) + 1);
  return counts;
};

/** Unique hints anchor the sequence. Repeated values never borrow a different native identity. */
export function alignTargetSequence(
  before: readonly string[],
  after: readonly string[]
): { readonly success: true; readonly alignment: TargetSequenceAlignment } | NativeEditFailure {
  const fail = (error: string): NativeEditFailure => ({
    success: false,
    code: 'unsupported-operation',
    error,
  });
  if (Math.max(before.length, after.length) > SIBLING_BLOCK_LIMITS.maxSiblings)
    return fail('Table sequence matching exceeds its sibling limit.');
  const oldCounts = targetKeyCounts(before),
    newCounts = targetKeyCounts(after),
    changes = diffArrays([...before], [...after], {
      timeout: SIBLING_BLOCK_LIMITS.diffTimeoutMs,
      comparator: (a, b) => a === b && oldCounts.get(a) === 1 && newCounts.get(b) === 1,
    });
  if (!changes) return fail('Table sequence matching exceeded its planning budget.');
  const pairs: { before: number; after: number }[] = [],
    gaps: { beforeStart: number; beforeEnd: number; afterStart: number; afterEnd: number }[] = [];
  const gap = (a: number, endA: number, b: number, endB: number) => {
    while (a < endA && b < endB) pairs.push({ before: a++, after: b++ });
    if (a < endA || b < endB)
      gaps.push({ beforeStart: a, beforeEnd: endA, afterStart: b, afterEnd: endB });
  };
  let a = 0,
    b = 0,
    gapA = 0,
    gapB = 0;
  for (const change of changes) {
    if (change.added) b += change.count;
    else if (change.removed) a += change.count;
    else {
      gap(gapA, a, gapB, b);
      for (let index = 0; index < change.count; index++) pairs.push({ before: a++, after: b++ });
      gapA = a;
      gapB = b;
    }
  }
  gap(gapA, a, gapB, b);
  pairs.sort((x, y) => x.before - y.before);
  return {
    success: true,
    alignment: Object.freeze({
      pairs: Object.freeze(pairs.map((pair) => Object.freeze(pair))),
      gaps: Object.freeze(gaps.map((gap) => Object.freeze(gap))),
    }),
  };
}
