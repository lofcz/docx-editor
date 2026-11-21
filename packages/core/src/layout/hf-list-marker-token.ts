import type { HeaderFooterStoryLayout } from './hf-layout.ts';
import { framedTokenJoin } from './framed-token.ts';
import { forEachStoryParagraphFragment } from './semantic-record-queries.ts';

const storyListMarkerTokens = new WeakMap<HeaderFooterStoryLayout, string>();

/** Numbering edits can change a header's markers without changing its authored content or flow height.
 * Include markers inside text boxes, which also read the separate numbering part.
 */
export function storyListMarkerToken(story: HeaderFooterStoryLayout): string {
  const cached = storyListMarkerTokens.get(story);
  if (cached !== undefined) return cached;
  const tokens: string[] = [];
  forEachStoryParagraphFragment(story, (fragment) => {
    const marker = fragment.marker;
    if (marker) {
      // Length-framed at both levels: `marker.text` is expanded `w:lvlText` and `numFmt`
      // is read verbatim from the file, so any separator the content can contain lets two
      // different marker states concatenate to one token and reuse a header page showing
      // the old numbers.
      tokens.push(
        framedTokenJoin([
          fragment.paragraphId,
          marker.text,
          marker.numFmt,
          String(marker.level),
          marker.numId ?? '',
          marker.abstractNumId ?? '',
        ])
      );
    }
  });
  const token = tokens.length === 0 ? '' : `|list:${framedTokenJoin(tokens)}`;
  storyListMarkerTokens.set(story, token);
  return token;
}
