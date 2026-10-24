/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { Step, Token } from './paragraph-text-diff';
import type * as Y from 'yjs';
import { isParagraphBoundary } from './paragraph-text';

/** Preserve explicit positions while separating unobserved prefixes from first typing. */
export function paragraphInsertionPosition(
  text: Y.Text,
  before: readonly Token[],
  inserted: readonly Token[],
  fallback: number
): number {
  if (fallback !== 0) return fallback;
  let boundary = 0;
  let found = false;
  for (const op of text.toDelta() as { insert: unknown }[]) {
    if (isParagraphBoundary(op.insert)) {
      found = true;
      break;
    }
    boundary += typeof op.insert === 'string' ? op.insert.length : 1;
  }
  // Hidden relocated text is not the landing of a visible empty paragraph.
  const first = before[0]?.position ?? (found ? boundary + 1 : text.length);
  return found && inserted[0] && inlineControlOf(inserted[0]) && first > boundary
    ? boundary
    : first;
}

/** A registered inline control is a wrapper around each of its characters or embeds. */
export function inlineControlOf(token: Token): string | null {
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
  const source = inlineControlOf(first),
    target = inserted[0] && inlineControlOf(inserted[0]);
  if (
    source &&
    target &&
    source !== target &&
    removed.every((token) => inlineControlOf(token) === source) &&
    inserted.every((token) => inlineControlOf(token) === target)
  )
    return first.position! + 1;
  return first.position!;
}
