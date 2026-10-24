import { segmentGraphemes } from '@docx-editor.dev/core/layout';
import { DEFAULT_TARGET_PARSE_LIMITS } from '../projection/parse-target';
import type { CompiledInlineTarget } from './inline-target';
import { PARAGRAPH_TEXT_LIMITS } from './paragraph-text';
import { retainedAtomKey } from './retained-atom-gaps';

interface Unit {
  readonly key: string;
  readonly atom?: number;
  readonly image?: boolean;
  readonly source?: boolean;
}
interface Trace {
  readonly old: number;
  readonly next: number;
  readonly previous: Trace | null;
}
interface State {
  readonly i: number;
  readonly j: number;
  readonly cost: number;
  readonly removed: number;
  readonly estimate: number;
  readonly order: number;
  readonly trace: Trace | null;
}
function units(target: CompiledInlineTarget): Unit[] {
  let atom = 0;
  return target.tokens.flatMap((token): Unit[] =>
    token.kind === 'text'
      ? segmentGraphemes(token.value).map(({ text }) => ({ key: 'text:' + text }))
      : [
          {
            key: 'atom:' + retainedAtomKey(token.ast),
            atom: atom++,
            image: token.ast.type === 'image',
            source: token.ast.name === 'source',
          },
        ]
  );
}
const precedes = (a: State, b: State) =>
  a.estimate < b.estimate ||
  (a.estimate === b.estimate &&
    (a.removed < b.removed ||
      (a.removed === b.removed &&
        (a.i + a.j > b.i + b.j || (a.i + a.j === b.i + b.j && a.order < b.order)))));

/**
 * Minimum text-edit alignment with mandatory source atoms. New equal atoms have fresh identities;
 * surrounding text chooses where old identities stay. Source atoms remain mandatory unless explicit removal is enabled.
 * Equal-cost paths prefer fewer native removals.
 * Optional image and source-payload replacements each cost one edit when enabled.
 * Search shares the existing paragraph diff deadline and target node bound.
 */
export function alignRetainedAtoms(
  before: CompiledInlineTarget,
  after: CompiledInlineTarget,
  replaceImages = false,
  removeAtoms = false,
  replaceSources = false
): readonly number[] | null {
  const deadline = performance.now() + PARAGRAPH_TEXT_LIMITS.diffTimeoutMs;
  const a = units(before),
    b = units(after);
  const heap: State[] = [];
  const weight = a.filter((unit) => unit.atom !== undefined).length + 1;
  const seen = new Map<number, number>();
  let order = 0,
    exhausted = false;
  const push = (from: number, to: number, cost: number, trace: Trace | null, removed = 0) => {
    let i = from,
      j = to;
    // Equal ordinary text is safe to keep. Atom matches remain explicit search edges.
    while (i < a.length && j < b.length && a[i]!.atom === undefined && a[i]!.key === b[j]!.key) {
      i++;
      j++;
      if ((i & 1023) === 0 && performance.now() >= deadline) {
        exhausted = true;
        return;
      }
    }
    const key = i * (b.length + 1) + j;
    const rank = cost * weight + removed;
    if (rank >= (seen.get(key) ?? Infinity)) return;
    if (seen.size >= DEFAULT_TARGET_PARSE_LIMITS.maxNodes) {
      exhausted = true;
      return;
    }
    seen.set(key, rank);
    const state: State = {
      i,
      j,
      cost,
      removed,
      trace,
      order: order++,
      estimate: cost + Math.abs(a.length - i - (b.length - j)),
    };
    let at = heap.length;
    heap.push(state);
    while (at > 0) {
      const parent = (at - 1) >>> 1;
      if (!precedes(state, heap[parent]!)) break;
      heap[at] = heap[parent]!;
      at = parent;
    }
    heap[at] = state;
  };
  const pop = (): State => {
    const first = heap[0]!,
      last = heap.pop()!;
    if (heap.length) {
      let at = 0;
      while (at * 2 + 1 < heap.length) {
        let child = at * 2 + 1;
        if (child + 1 < heap.length && precedes(heap[child + 1]!, heap[child]!)) child++;
        if (!precedes(heap[child]!, last)) break;
        heap[at] = heap[child]!;
        at = child;
      }
      heap[at] = last;
    }
    return first;
  };
  push(0, 0, 0, null);
  while (heap.length && !exhausted && performance.now() < deadline) {
    const state = pop(),
      { i, j, cost, trace, removed } = state;
    if (seen.get(i * (b.length + 1) + j) !== cost * weight + removed) continue;
    if (i === a.length && j === b.length) {
      const matches: number[] = Array(a.filter((unit) => unit.atom !== undefined).length).fill(-1);
      for (let step = trace; step; step = step.previous) matches[step.old] = step.next;
      return matches;
    }
    const old = a[i],
      next = b[j];
    if (old?.atom !== undefined && next?.atom !== undefined) {
      const replacement =
        old.key === next.key
          ? 0
          : (replaceImages && old.image && next.image) ||
              (replaceSources && old.source && next.source)
            ? 1
            : null;
      if (replacement !== null)
        push(
          i + 1,
          j + 1,
          cost + replacement,
          { old: old.atom, next: next.atom, previous: trace },
          removed
        );
    }
    if (next) push(i, j + 1, cost + 1, trace, removed);
    if (old && (old.atom === undefined || removeAtoms))
      push(i + 1, j, cost + 1, trace, removed + (old.atom === undefined ? 0 : 1));
  }
  return null;
}
