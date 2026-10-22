import type { OoxmlNode } from '../package/ooxml-tree.ts';
import { transactionActorId } from '../package/actor-scoped-ids.ts';
import { retireParagraph } from '../package/retired-paragraph.ts';
import { mergeMarkedParagraphs } from './paragraph-mark-merge.ts';

export { paragraphMergeSources, tableParagraphMergeTarget } from './paragraph-mark-merge.ts';

/** Keep actor-bound source identities; Word export alone performs the physical join. */
export function mergeRevisionParagraphs(
  children: OoxmlNode[],
  marks: ReadonlySet<string>,
  mint: () => string
): OoxmlNode[] {
  return mergeMarkedParagraphs(
    children,
    marks,
    mint,
    transactionActorId() === undefined
      ? undefined
      : (paragraph, forward) => retireParagraph(paragraph, mint, forward)
  );
}
