import type { OoxmlParagraphNode } from '../package/ooxml-tree.ts';
import { retractedLengthOf } from './tree-op-retraction.ts';
import { trackedInsertionLanding } from './tree-op-tracked-adjacency.ts';

/**
 * Where a tracked replacement of `[start, end)` in `paragraph` lands, from the paragraph's own
 * tree: after the struck words, minus what of the range was the author's own pending insertion,
 * which retracts. The store's rule, so a caller cannot predict a landing the store will not use.
 *
 * This pure query does not enable tracking or apply operations.
 * @public
 */
export function trackedReplacementLanding(
  paragraph: OoxmlParagraphNode,
  start: number,
  end: number,
  author: string,
  date: string | undefined
): number {
  const { past, landing } = trackedInsertionLanding(paragraph, { start, end }, end, author, date);
  const retracted = retractedLengthOf(paragraph, start, past, author);
  // NOTHING STRUCK, nothing to land after: a zero-width range, or one covering only this
  // author's own pending insertion, which retracts and writes no strike at all. The store
  // then places at the aim, and so must this.
  const strikes = past - start > retracted;
  return (strikes ? landing : past) - retracted;
}
