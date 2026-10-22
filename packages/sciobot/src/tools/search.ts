import type { SearchError, SearchInput, SearchMatch, SearchResult } from '../contracts/executor';
import type { MarkdownProjection } from '../projection/markdown';

/** Preserve the production search contract: one result per matching line. */
export function searchDocument(
  projection: Pick<MarkdownProjection, 'text'>,
  input: Readonly<SearchInput>
): SearchResult | SearchError {
  if (input.pattern.length === 0) {
    return { matches: [], totalMatches: 0, error: 'pattern must be a non-empty string' };
  }
  let regex: RegExp;
  try {
    regex = input.isRegex
      ? new RegExp(input.pattern, 'gim')
      : new RegExp(input.pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gim');
  } catch (cause) {
    return {
      matches: [],
      totalMatches: 0,
      error: `Invalid regex: ${cause instanceof Error ? cause.message : String(cause)}`,
    };
  }
  const lines = projection.text.split('\n');
  const matches: SearchMatch[] = [];
  let matchCount = 0;
  for (let index = 0; index < lines.length; index += 1) {
    if (regex.test(lines[index]!)) {
      matchCount += 1;
      if (matches.length < input.maxResults)
        matches.push({
          lineNumber: index + 1,
          content: lines[index]!,
          contextBefore: index > 0 ? lines[index - 1]! : null,
          contextAfter: index < lines.length - 1 ? lines[index + 1]! : null,
        });
      regex.lastIndex = 0;
    }
  }
  return { matches, totalMatches: matchCount, truncated: matchCount > input.maxResults };
}
