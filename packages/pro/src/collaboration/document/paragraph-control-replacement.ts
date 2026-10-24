/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { Step, Token } from './paragraph-text-diff';

/** A registered inline control is a wrapper around each of its characters or embeds. */
function control(token: Token): string | null {
  try {
    const wrappers: unknown = JSON.parse(token.attributes.w ?? 'null');
    if (!Array.isArray(wrappers)) return null;
    const wrapper = wrappers.find((entry) => entry?.k === 'contentControl' && entry?.l === 'sdt');
    return typeof wrapper?.i === 'string' ? wrapper.i : null;
  } catch {
    return null;
  }
}

/**
 * Replacing a whole inline control has an interior native boundary. Anchor its new
 * content after the first removed character, so typing before the old control stays
 * before it and typing after a multi-character control stays after it. Ordinary text
 * replacement and edits within the same control retain their established boundary.
 */
export function controlReplacementPosition(
  before: readonly Token[],
  after: readonly Token[],
  steps: readonly Step[],
  fallback: number
): number {
  const removed = steps.flatMap((step) => (step.op === 'del' ? [before[step.before]!] : []));
  const inserted = steps.flatMap((step) => (step.op === 'ins' ? [after[step.after]!] : []));
  const first = removed[0];
  if (!first) return fallback;
  const source = control(first),
    target = inserted[0] && control(inserted[0]);
  if (
    source &&
    target &&
    source !== target &&
    removed.every((token) => control(token) === source) &&
    inserted.every((token) => control(token) === target)
  )
    return first.position! + 1;
  return first.position!;
}
