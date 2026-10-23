import { logicalParagraphSlice } from './joined-paragraphs.ts';
import { trimmed } from './plan-read-helpers.ts';
import type { AutomationStoryReads } from './reads.ts';
import type { TreeDocOp } from '../store/store/tree-ops.ts';

/** Native edits after a logical paragraph has been consolidated into its first fragment. */
export function paragraphSplitEdits(
  paragraphId: string,
  text: string,
  occurrences: readonly { readonly start: number; readonly length: number }[],
  trimDelimiters: boolean
) {
  const ops: TreeDocOp[] = [];
  const offsets: number[] = [];
  /** The text each resulting paragraph will hold, for the answered ranges. */
  const pieces: string[] = [];

  if (occurrences.length === 0) {
    pieces.push(text);
  } else if (trimDelimiters) {
    // Cut the delimiters out from the LAST backwards, so every remaining op's offsets still
    // describe the paragraph the caller measured.
    for (const occurrence of [...occurrences].reverse()) {
      ops.push({
        op: 'deleteText',
        paragraphId: paragraphId,
        start: occurrence.start,
        end: occurrence.start + occurrence.length,
      });
    }
    let removed = 0;
    let previous = 0;
    for (const occurrence of occurrences) {
      const at = occurrence.start - removed;
      offsets.push(at);
      pieces.push(text.slice(previous, occurrence.start));
      previous = occurrence.start + occurrence.length;
      removed += occurrence.length;
    }
    pieces.push(text.slice(previous));
  } else {
    let previous = 0;
    for (const occurrence of occurrences) {
      const at = occurrence.start + occurrence.length;
      offsets.push(at);
      pieces.push(text.slice(previous, at));
      previous = at;
    }
    pieces.push(text.slice(previous));
  }

  return { ops, offsets, pieces };
}

/** Keep the split admission limits shared by physical and logical paragraphs. */
export function invalidSplitDelimiters(
  delimiters: unknown
): { message: string; detail: string } | null {
  if (!Array.isArray(delimiters) || delimiters.length === 0)
    return { message: 'split needs at least one delimiter', detail: 'delimiters' };
  if (delimiters.length > 16)
    return { message: 'too many delimiters', detail: String(delimiters.length) };
  for (const delimiter of delimiters) {
    if (typeof delimiter !== 'string' || delimiter.length === 0)
      return { message: 'a delimiter must be a non-empty string', detail: 'delimiters' };
    if (delimiter.length > 64)
      return { message: 'that delimiter is too long', detail: 'delimiters' };
  }
  return null;
}

export function unchangedLogicalSplitRange(
  reads: AutomationStoryReads,
  id: string,
  text: string,
  trimSpacing: boolean
) {
  const [from, to] = trimSpacing ? trimmed(text, 0, text.length) : [0, text.length];
  return logicalParagraphSlice(reads, id, from!, to!);
}
