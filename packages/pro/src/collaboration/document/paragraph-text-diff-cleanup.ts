/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { Step, Token } from './paragraph-text-diff.ts';

/**
 * Keep the letters `dropScatteredLetters` wrote again in a change that types nothing.
 *
 * Only new text can take the place of a kept letter. A change whose inserts are all letters
 * written again only deletes, as the head of a paragraph does when Enter moves its tail away
 * and a second copy of an embed before its last letter goes too. Written again there, a letter
 * is a new character, and a peer's concurrent deletion of the old one does not reach it, so
 * the deleted text shows again. A change is a run of steps between kept characters, so letters
 * the cleanup merged into a replacement that types new text stay replaced.
 */
export function keepLettersTypedOverByNothing(
  steps: Step[],
  rewritten: ReadonlyMap<number, number>
): Step[] {
  if (rewritten.size === 0) return steps;
  const keptAt = new Map<number, number>();
  for (const [after, before] of rewritten) keptAt.set(before, after);
  const out: Step[] = [];
  for (let at = 0; at < steps.length;) {
    const step = steps[at]!;
    if (step.op === 'eq') {
      out.push(step);
      at += 1;
      continue;
    }
    let end = at;
    while (end < steps.length && steps[end]!.op !== 'eq') end += 1;
    const change = steps.slice(at, end);
    at = end;
    if (change.some((entry) => entry.op === 'ins' && !rewritten.has(entry.after))) {
      out.push(...change);
      continue;
    }
    // Each letter written again is deleted in its place in `before` order: keep it there.
    for (const entry of change) {
      if (entry.op !== 'del') continue;
      const after = keptAt.get(entry.before);
      out.push(after === undefined ? entry : { op: 'eq', before: entry.before, after });
    }
  }
  return out;
}

/**
 * Join inserts that a kept character splits, when the second insert ends with that same
 * letter: keep the later copy instead, and the inserted text is one piece, as it was typed.
 * Both scripts keep as many characters; only which of two equal letters stays differs.
 */
export function joinInserts(steps: Step[], after: readonly Token[]): Step[] {
  for (let at = 1; at < steps.length; at += 1) {
    const kept = steps[at]!;
    if (kept.op !== 'eq' || steps[at - 1]!.op !== 'ins') continue;
    let end = at + 1;
    while (end < steps.length && steps[end]!.op === 'ins') end += 1;
    const last = steps[end - 1]!;
    if (end === at + 1 || last.op !== 'ins') continue;
    if (after[last.after]!.key !== after[kept.after]!.key) continue;
    steps.splice(at, end - at, { op: 'ins', after: kept.after }, ...steps.slice(at + 1, end - 1), {
      op: 'eq',
      before: kept.before,
      after: last.after,
    });
  }
  return steps;
}
