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

export function unchangedLogicalSplitRange(
  reads: AutomationStoryReads,
  id: string,
  text: string,
  trimSpacing: boolean
) {
  const [from, to] = trimSpacing ? trimmed(text, 0, text.length) : [0, text.length];
  return logicalParagraphSlice(reads, id, from!, to!);
}
