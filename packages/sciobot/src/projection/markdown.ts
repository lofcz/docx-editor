import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import {
  exportMarkdownLayout,
  type MarkdownSourceMap,
  type MarkdownSourceSlice,
  type MarkdownLayoutOptions,
} from '@docx-editor.dev/docx-to-markdown';
import type { RevisionToken } from '../contracts/document';
import { SCIOBOT_CODE_CHARACTER_STYLES } from '../native/inline-code';
import type { NativeTextPoint } from './text';

export interface NativeMarkdownPoint extends NativeTextPoint {
  readonly sourceScope: string;
}

export interface MarkdownSyntaxSpan {
  readonly start: number;
  readonly end: number;
  readonly role: 'generated';
}

/** Immutable read data. The caller supplies the layout for the required review view. */
export interface MarkdownProjection {
  readonly revision: RevisionToken;
  readonly displayMode: ExportSemanticLayout['displayMode'];
  readonly text: string;
  readonly sourceMap: MarkdownSourceMap;
  readonly syntax: readonly MarkdownSyntaxSpan[];
}

/** Read native logical records. Do not deserialize Markdown back over the imported document. */
export function projectMarkdown(
  layout: ExportSemanticLayout,
  revision: RevisionToken,
  options: Pick<MarkdownLayoutOptions, 'tableSyntax' | 'codeCharacterStyles'> = {}
): MarkdownProjection {
  // Code spans come only from the explicit application character style, never from fonts.
  const result = exportMarkdownLayout(layout, {
    codeCharacterStyles: SCIOBOT_CODE_CHARACTER_STYLES,
    ...options,
    sourceMap: true,
  });
  const nativeSourceMap = result.sourceMap!;
  // Native optional fields can be explicitly undefined. Keep this private read DTO plain JSON.
  const sources = nativeSourceMap.sources.map((source) => {
    if (source.markdownBoundaries !== undefined || !Object.hasOwn(source, 'markdownBoundaries'))
      return source;
    const normalized = { ...source };
    delete normalized.markdownBoundaries;
    return Object.freeze(normalized);
  });
  const sourceMap = sources.every((source, index) => source === nativeSourceMap.sources[index])
    ? nativeSourceMap
    : Object.freeze({ ...nativeSourceMap, sources: Object.freeze(sources) });
  const syntax: MarkdownSyntaxSpan[] = [];
  let cursor = 0;
  for (const slice of sourceMap.sources) {
    if (cursor < slice.markdownStart)
      syntax.push(Object.freeze({ start: cursor, end: slice.markdownStart, role: 'generated' }));
    cursor = Math.max(cursor, slice.markdownEnd);
  }
  if (cursor < result.markdown.length)
    syntax.push(Object.freeze({ start: cursor, end: result.markdown.length, role: 'generated' }));
  return Object.freeze({
    revision: Object.freeze({ ...revision }),
    displayMode: layout.displayMode,
    text: result.markdown,
    sourceMap,
    syntax: Object.freeze(syntax),
  });
}

export function sourceOffsetsAt(slice: MarkdownSourceSlice, offset: number): readonly number[] {
  if (!slice.exact || offset < slice.markdownStart || offset > slice.markdownEnd) return [];
  const boundaries = slice.markdownBoundaries ?? [];
  let low = 0;
  let high = boundaries.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (boundaries[middle]!.markdownOffset < offset) low = middle + 1;
    else high = middle;
  }
  const boundary = boundaries[low];
  if (boundary?.markdownOffset === offset) {
    const found: number[] = [];
    while (low < boundaries.length && boundaries[low]!.markdownOffset === offset)
      found.push(boundaries[low++]!.sourceOffset);
    return found;
  }
  const before = boundaries[low - 1];
  if (
    !before ||
    !boundary ||
    boundary.markdownOffset - before.markdownOffset !== boundary.sourceOffset - before.sourceOffset
  )
    return [];
  return [before.sourceOffset + offset - before.markdownOffset];
}

/**
 * Escape interiors, generated delimiters, and surrogate splits fail closed.
 * Hidden review text can give one visible boundary two native positions. Select an explicit
 * affinity for an edit endpoint, or omit it to reject that ambiguity.
 */
export function markdownPointAt(
  projection: MarkdownProjection,
  offset: number,
  affinity?: 'forward' | 'backward'
): NativeMarkdownPoint | null {
  if (!Number.isInteger(offset) || offset < 0 || offset > projection.text.length) return null;
  const before = projection.text.charCodeAt(offset - 1);
  const after = projection.text.charCodeAt(offset);
  if (before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff) return null;
  let point: NativeMarkdownPoint | null = null;
  for (const slice of projection.sourceMap.sources) {
    const offsets = sourceOffsetsAt(slice, offset);
    if (new Set(offsets).size > 1) return null;
    for (const nativeOffset of offsets) {
      const candidate = {
        sourceScope: slice.sourceScope,
        paragraphId: slice.paragraphId,
        offset: nativeOffset,
      };
      if (
        point &&
        (point.sourceScope !== candidate.sourceScope ||
          point.paragraphId !== candidate.paragraphId ||
          point.offset !== candidate.offset)
      ) {
        if (!affinity || point.sourceScope !== candidate.sourceScope) return null;
        if (affinity === 'forward') point = candidate;
      } else point = candidate;
    }
  }
  return point;
}

/** This resolves text endpoints only. P10 supplies structural and protected-content checks. */
export function markdownRangeAt(
  projection: MarkdownProjection,
  start: number,
  end: number
): { readonly start: NativeMarkdownPoint; readonly end: NativeMarkdownPoint } | null {
  if (end < start) return null;
  const from = markdownPointAt(projection, start, 'forward');
  const to = markdownPointAt(projection, end, start === end ? 'forward' : 'backward');
  return from && to && from.sourceScope === to.sourceScope ? { start: from, end: to } : null;
}
