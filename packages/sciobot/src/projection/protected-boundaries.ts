import type {
  MarkdownSourceParagraph,
  MarkdownSourceSlice,
} from '@docx-editor.dev/docx-to-markdown';
import type { MarkdownProjection } from './markdown';
import type { ProtectedMarkdownSpan } from './resolve-patch';

export interface ProtectedNativeRange {
  readonly paragraphId: string;
  readonly start: number;
  readonly end: number;
}
export interface ProtectedContentBoundary extends ProtectedMarkdownSpan {
  readonly partName: string;
  readonly sourceScope: string;
  readonly native?: ProtectedNativeRange;
}
interface Extent {
  start: number;
  end: number;
}
function point(slice: MarkdownSourceSlice, offset: number): number | null {
  if (offset === slice.sourceStart) return slice.markdownStart;
  if (offset === slice.sourceEnd) return slice.markdownEnd;
  const boundaries = slice.markdownBoundaries ?? [];
  for (let index = 0; index < boundaries.length; index++) {
    const a = boundaries[index]!,
      b = boundaries[index + 1];
    if (a.sourceOffset === offset) return a.markdownOffset;
    if (
      b &&
      a.sourceOffset < offset &&
      offset < b.sourceOffset &&
      b.sourceOffset - a.sourceOffset === b.markdownOffset - a.markdownOffset
    )
      return a.markdownOffset + offset - a.sourceOffset;
  }
  return null;
}

/** Native coordinates come only from core's offset index. No displayed text matching occurs. */
export function protectedRangeBoundaries(
  projection: MarkdownProjection,
  range: ProtectedNativeRange,
  partName: string,
  nodeId: string,
  reason: string
): ProtectedContentBoundary[] {
  const paragraphs = projection.sourceMap.paragraphs.filter(
    (p) => p.sourceScope === 'body' && p.paragraphId === range.paragraphId
  );
  const out: ProtectedContentBoundary[] = [];
  for (const paragraph of paragraphs) {
    const sources = projection.sourceMap.sources.filter(
      (source) => source.sourceScope === 'body' && source.paragraphId === range.paragraphId
    );
    const extents: Extent[] = [];
    if (range.start < range.end)
      for (const slice of sources) {
        const start = Math.max(range.start, slice.sourceStart),
          end = Math.min(range.end, slice.sourceEnd);
        if (start >= end) continue;
        const a = slice.exact ? point(slice, start) : null,
          b = slice.exact ? point(slice, end) : null;
        extents.push({ start: a ?? slice.markdownStart, end: b ?? slice.markdownEnd });
      }
    if (!extents.length) {
      const locations = sources.flatMap((source) => {
        if (
          range.start === range.end &&
          source.sourceStart <= range.start &&
          range.end <= source.sourceEnd
        ) {
          const at = point(source, range.start);
          return at === null ? [] : [at];
        }
        return [];
      });
      if (locations.length)
        extents.push({
          start: locations.reduce((at, value) => Math.min(at, value), Infinity),
          end: locations.reduce((at, value) => Math.max(at, value), -Infinity),
        });
      else {
        const left = sources
          .filter((source) => source.sourceEnd <= range.start)
          .sort((a, b) => b.sourceEnd - a.sourceEnd)[0];
        const right = sources
          .filter((source) => source.sourceStart >= range.end)
          .sort((a, b) => a.sourceStart - b.sourceStart)[0];
        const start = left?.markdownEnd ?? paragraph.markdownStart;
        const end = right?.markdownStart ?? paragraph.markdownEnd;
        if (start !== undefined && end !== undefined && start <= end) extents.push({ start, end });
      }
    }
    if (!extents.length) continue;
    const start = extents.reduce((value, item) => Math.min(value, item.start), Infinity);
    const end = extents.reduce((value, item) => Math.max(value, item.end), -Infinity);
    out.push(
      Object.freeze({
        start,
        end,
        partName,
        nodeId,
        reason,
        sourceScope: 'body',
        ...(start === end ? { edges: 'interior' as const } : {}),
        native: Object.freeze({ ...range }),
      })
    );
  }
  return out;
}
export function outsideBodyBoundary(
  paragraph: MarkdownSourceParagraph,
  partName: string
): ProtectedContentBoundary | null {
  if (paragraph.markdownStart === undefined || paragraph.markdownEnd === undefined) return null;
  return Object.freeze({
    start: paragraph.markdownStart,
    end: paragraph.markdownEnd,
    partName,
    nodeId: paragraph.paragraphId,
    sourceScope: paragraph.sourceScope,
    reason: `Content outside the writable body (${paragraph.sourceScope})`,
  });
}
