import type {
  MarkdownSourceParagraph,
  MarkdownSourceSlice,
  MarkdownSourceTable,
} from '@docx-editor.dev/docx-to-markdown';
import { sourceOffsetsAt, type MarkdownProjection, type MarkdownSyntaxSpan } from './markdown';

interface Piece {
  readonly kind: 'native' | 'replacement';
  readonly start: number;
  readonly end: number;
  readonly outputStart: number;
  readonly outputEnd: number;
}

/** Builds derived text only. No canonical nodes or editor operations enter this writer. */
export class ProjectionComposer {
  private readonly chunks: string[] = [];
  private readonly pieces: Piece[] = [];
  private length = 0;

  constructor(private readonly base: MarkdownProjection) {}

  get offset(): number {
    return this.length;
  }
  get endsWithParagraphBreak(): boolean {
    let tail = '';
    for (let at = this.chunks.length - 1; at >= 0 && tail.length < 2; at--)
      tail = this.chunks[at]!.slice(-2) + tail;
    return tail.endsWith('\n\n');
  }

  generated(text: string): void {
    this.chunks.push(text);
    this.length += text.length;
  }

  native(start: number, end: number): void {
    const outputStart = this.length;
    this.generated(this.base.text.slice(start, end));
    this.pieces.push({ kind: 'native', start, end, outputStart, outputEnd: this.length });
  }

  replacement(start: number, end: number, text: string): void {
    const outputStart = this.length;
    this.generated(text);
    this.pieces.push({ kind: 'replacement', start, end, outputStart, outputEnd: this.length });
  }

  private firstPiece(offset: number): number {
    let low = 0,
      high = this.pieces.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (this.pieces[middle]!.end < offset) low = middle + 1;
      else high = middle;
    }
    return low;
  }

  mapExtent(
    start: number,
    end: number,
    includeStartInsertions = true
  ): { markdownStart: number; markdownEnd: number } | null {
    if (start === end) {
      for (let index = this.firstPiece(start); index < this.pieces.length; index += 1) {
        const piece = this.pieces[index]!;
        if (piece.start > start) break;
        if (piece.start === start && piece.end === end)
          return { markdownStart: piece.outputStart, markdownEnd: piece.outputEnd };
      }
    }
    let markdownStart = Infinity,
      markdownEnd = -Infinity;
    for (let index = this.firstPiece(start); index < this.pieces.length; index += 1) {
      const piece = this.pieces[index]!;
      if (piece.start > end) break;
      if (start !== end && piece.start === piece.end && piece.start === end) continue;
      if (
        start !== end &&
        !includeStartInsertions &&
        piece.start === piece.end &&
        piece.start === start
      )
        continue;
      const from = Math.max(start, piece.start),
        to = Math.min(end, piece.end);
      if (from > to || (from === to && start !== end && piece.kind === 'native')) continue;
      const a =
        piece.kind === 'replacement' ? piece.outputStart : piece.outputStart + from - piece.start;
      const b =
        piece.kind === 'replacement' ? piece.outputEnd : piece.outputStart + to - piece.start;
      markdownStart = Math.min(markdownStart, a);
      markdownEnd = Math.max(markdownEnd, b);
    }
    return Number.isFinite(markdownStart) ? { markdownStart, markdownEnd } : null;
  }

  finish(): MarkdownProjection {
    const sources: MarkdownSourceSlice[] = [];
    for (const source of this.base.sourceMap.sources) {
      for (
        let index = this.firstPiece(source.markdownStart);
        index < this.pieces.length;
        index += 1
      ) {
        const piece = this.pieces[index]!;
        if (piece.start >= source.markdownEnd) break;
        if (piece.kind !== 'native') continue;
        const start = Math.max(piece.start, source.markdownStart);
        const end = Math.min(piece.end, source.markdownEnd);
        if (start >= end) continue;
        const delta = piece.outputStart - piece.start;
        let sourceStart = source.sourceStart,
          sourceEnd = source.sourceEnd;
        if (start !== source.markdownStart || end !== source.markdownEnd) {
          const starts = new Set(sourceOffsetsAt(source, start));
          const ends = new Set(sourceOffsetsAt(source, end));
          if (starts.size !== 1 || ends.size !== 1)
            throw new Error('A block boundary has no exact native text position.');
          sourceStart = starts.values().next().value!;
          sourceEnd = ends.values().next().value!;
        }
        const boundaries = source.markdownBoundaries
          ? [
              { sourceOffset: sourceStart, markdownOffset: start + delta },
              ...source.markdownBoundaries
                .filter(
                  (boundary) => start < boundary.markdownOffset && boundary.markdownOffset < end
                )
                .map((boundary) => ({
                  ...boundary,
                  markdownOffset: boundary.markdownOffset + delta,
                })),
              { sourceOffset: sourceEnd, markdownOffset: end + delta },
            ]
          : undefined;
        sources.push(
          Object.freeze({
            ...source,
            sourceStart,
            sourceEnd,
            markdownStart: start + delta,
            markdownEnd: end + delta,
            ...(boundaries
              ? {
                  markdownBoundaries: Object.freeze(
                    boundaries.map((boundary) => Object.freeze(boundary))
                  ),
                }
              : {}),
          })
        );
      }
    }
    sources.sort(
      (left, right) =>
        left.markdownStart - right.markdownStart || left.markdownEnd - right.markdownEnd
    );
    const paragraphs: MarkdownSourceParagraph[] = [];
    for (const paragraph of this.base.sourceMap.paragraphs) {
      if (paragraph.markdownStart === undefined || paragraph.markdownEnd === undefined) continue;
      const extent = this.mapExtent(paragraph.markdownStart, paragraph.markdownEnd);
      if (extent) paragraphs.push(Object.freeze({ ...paragraph, ...extent }));
    }
    const tables: MarkdownSourceTable[] = [];
    for (const table of this.base.sourceMap.tables) {
      const extent = this.mapExtent(table.markdownStart, table.markdownEnd);
      if (extent) tables.push(Object.freeze({ ...table, ...extent }));
    }
    const text = this.chunks.join('');
    const syntax: MarkdownSyntaxSpan[] = [];
    let cursor = 0;
    for (const source of sources) {
      if (cursor < source.markdownStart)
        syntax.push(Object.freeze({ start: cursor, end: source.markdownStart, role: 'generated' }));
      cursor = Math.max(cursor, source.markdownEnd);
    }
    if (cursor < text.length)
      syntax.push(Object.freeze({ start: cursor, end: text.length, role: 'generated' }));
    return Object.freeze({
      ...this.base,
      text,
      revision: Object.freeze({ ...this.base.revision }),
      sourceMap: Object.freeze({
        unit: 'utf16-code-unit',
        sources: Object.freeze(sources),
        paragraphs: Object.freeze(paragraphs),
        tables: Object.freeze(tables),
      }),
      syntax: Object.freeze(syntax),
    });
  }
}
