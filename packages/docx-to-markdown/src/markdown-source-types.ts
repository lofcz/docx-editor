/** Source boundaries for the immutable logical `markdown` string. @public */
export interface MarkdownSourceMap {
  /** Coordinate unit for both native text and generated Markdown. */
  readonly unit: 'utf16-code-unit';
  /** Text ranges exclude generated headings, emphasis delimiters, and table separators. */
  readonly sources: readonly MarkdownSourceSlice[];
  /** Source extents include empty paragraphs that emit no text. */
  readonly paragraphs: readonly MarkdownSourceParagraph[];
  /** Complete table constructs, including generated GFM separators or nested HTML wrappers. */
  readonly tables: readonly MarkdownSourceTable[];
}

/** A complete logical table construct in the generated Markdown string. @public */
export interface MarkdownSourceTable {
  /** `body`, or a story kind followed by its part or note scope identifier. */
  readonly sourceScope: string;
  /** Canonical `w:tbl` node ID. Physical page fragments share this identity. */
  readonly tableId: string;
  /** Inclusive construct position, including its generated opening syntax. */
  readonly markdownStart: number;
  /** Exclusive construct position, including its generated closing syntax. */
  readonly markdownEnd: number;
}

/** One source-backed substring inside the logical Markdown projection. @public */
export interface MarkdownSourceSlice {
  /** `body`, or a story kind followed by its part or note scope identifier. */
  readonly sourceScope: string;
  /** Canonical paragraph node ID within the source story. */
  readonly paragraphId: string;
  /** Inclusive source position within the native paragraph. */
  readonly sourceStart: number;
  /** Exclusive source position within the native paragraph. */
  readonly sourceEnd: number;
  /** Inclusive position within the logical Markdown string. */
  readonly markdownStart: number;
  /** Exclusive position within the logical Markdown string. */
  readonly markdownEnd: number;
  /**
   * Ordered absolute boundaries. Interpolate only when both coordinate deltas are equal.
   * Other gaps contain escaping or generated syntax and have no interior text boundary.
   * UTF-16 coordinates can bisect surrogate pairs; callers must reject such edit positions.
   */
  readonly markdownBoundaries?: readonly MarkdownSourceBoundary[];
  /** False for generated fallbacks without exact source character boundaries. */
  readonly exact: boolean;
}

/** Full source extent of one paragraph, including paragraphs with no Markdown text. @public */
export interface MarkdownSourceParagraph {
  /** `body`, or a story kind followed by its part or note scope identifier. */
  readonly sourceScope: string;
  /** Canonical paragraph node ID within the source story. */
  readonly paragraphId: string;
  /** Inclusive source position in this paragraph's represented extent. */
  readonly sourceStart: number;
  /** Exclusive source position in this paragraph's represented extent. */
  readonly sourceEnd: number;
  /** Inclusive logical Markdown construct position, including generated paragraph syntax. */
  readonly markdownStart?: number;
  /** Exclusive construct position. Empty paragraphs can have equal start and end positions. */
  readonly markdownEnd?: number;
}

/** A boundary in both the native paragraph and logical Markdown string. @public */
export interface MarkdownSourceBoundary {
  /** Position within the native source paragraph. */
  readonly sourceOffset: number;
  /** Corresponding position within the logical Markdown string. */
  readonly markdownOffset: number;
}

/** Source mapping controls shared by live, detached, and one-shot exports. @public */
export interface MarkdownSourceMapOptions {
  /** Include a read-only logical source map. Default: false. No document editing is performed. */
  readonly sourceMap?: boolean;
}
