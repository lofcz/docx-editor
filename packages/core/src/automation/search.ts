import {
  logicalParagraphIds,
  logicalParagraphPoint,
  logicalParagraphOffset,
  logicalParagraphProjection,
} from './joined-paragraphs.ts';
// Projected text search with model-offset results.

import { searchableLineBreaks } from './line-break-text.ts';
import { SEARCH_MATCH_LIMIT } from '../store/store/text-match.ts';
import type { AutomationHandleTable } from './handles.ts';
import type { AutomationSearchOptions } from './operations.ts';
import type { AutomationSpan } from './protocol.ts';
import type { AutomationStoryReads } from './reads.ts';
import { spanParagraphIds, spanValue, type ResolvedSpan } from './spans.ts';

export type ProjectedSearchResult =
  | { readonly ok: true; readonly spans: readonly AutomationSpan[] }
  | { readonly ok: false; readonly paragraphId: string };

/** Search one story view and map each visible hit back to one editable model span. */
export function projectedSearchSpans(
  reads: AutomationStoryReads,
  scope: ResolvedSpan,
  handles: AutomationHandleTable,
  text: string,
  options: AutomationSearchOptions | undefined
): ProjectedSearchResult {
  let budget = Math.min(options?.limit ?? SEARCH_MATCH_LIMIT, SEARCH_MATCH_LIMIT);
  const spans: AutomationSpan[] = [];
  const ids = logicalParagraphIds(reads, spanParagraphIds(scope, reads));
  const last = ids.length - 1;
  for (const [position, paragraphId] of ids.entries()) {
    if (budget <= 0) break;
    const projected = logicalParagraphProjection(
      reads,
      paragraphId,
      options?.projection ?? 'allMarkup'
    );
    if (!projected) continue;
    // Text reads report a line break as `\v`. A query that still spells it `\n` finds it too.
    const found = projected.findOccurrences(searchableLineBreaks(text), budget, {
      matchCase: options?.matchCase === true,
      wholeWord: options?.matchWholeWord === true,
      ...(position === 0 && scope
        ? {
            from: projected.projectedOffset(
              logicalParagraphOffset(reads, scope.start.paragraphId, scope.start.offset)
            ),
          }
        : {}),
      ...(position === last && scope
        ? {
            to: projected.projectedOffset(
              logicalParagraphOffset(reads, scope.end.paragraphId, scope.end.offset)
            ),
          }
        : {}),
    });
    for (const occurrence of found.matches) {
      const start = logicalParagraphPoint(reads, paragraphId, occurrence.rawStart);
      const end = logicalParagraphPoint(reads, paragraphId, occurrence.rawEnd, 'end');
      if (!start || !end) return { ok: false, paragraphId };
      spans.push(spanValue({ start, end }, handles));
      budget -= 1;
    }
  }
  return { ok: true, spans };
}
