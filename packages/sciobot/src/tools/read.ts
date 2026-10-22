import type { ReadInput, ReadResult } from '../contracts/executor';
import type { MarkdownProjection } from '../projection/markdown';

/** Effective client behavior. Server prompt guidance separately declares 500 lines. */
export const READ_DOCUMENT_DEFAULT_LIMIT = 1000;

/** Return copyable Markdown without line-number prefixes or projection mutations. */
export function readDocument(
  projection: Pick<MarkdownProjection, 'text'>,
  input: Readonly<ReadInput> = {}
): ReadResult {
  const lines = projection.text.split('\n');
  const totalLines = lines.length;
  const offsetRaw = input.offset;
  const limitRaw = input.limit;
  const offset =
    typeof offsetRaw === 'number' && Number.isFinite(offsetRaw) && offsetRaw >= 1
      ? Math.floor(offsetRaw)
      : 1;
  const limit =
    typeof limitRaw === 'number' && Number.isFinite(limitRaw) && limitRaw >= 1
      ? Math.floor(limitRaw)
      : READ_DOCUMENT_DEFAULT_LIMIT;
  const startIdx = Math.max(0, offset - 1);
  const endIdx = Math.min(totalLines, startIdx + limit);
  return {
    content: lines.slice(startIdx, endIdx).join('\n'),
    startLine: startIdx + 1,
    endLine: endIdx,
    totalLines,
    truncated: endIdx < totalLines,
  };
}
