import {
  revisionsAreDeletion,
  type SemanticLayout,
  type StyleSpanRecord,
} from '@docx-editor.dev/core/layout';
import {
  concatMarkdown,
  literalMarkdown,
  type MappedMarkdown,
  type MarkdownSourceSlice,
} from './markdown-source-map.ts';

// `code` is always innermost and literal: its content is never escaped and it opens with a
// content-dependent backtick fence rather than a fixed delimiter.
type MarkdownInlineStyle = 'strike' | 'italic' | 'bold' | 'code' | 'codeHtml';

export interface MarkdownTextToken {
  readonly span: StyleSpanRecord;
  readonly paragraphId: string;
  readonly sourceText: string;
  readonly sourceOffset?: number;
  readonly exact?: boolean;
}

type MarkdownInlineChunk =
  | {
      readonly kind: 'text';
      readonly sourceText: string;
      readonly styles: readonly MarkdownInlineStyle[];
      readonly semanticStyles: readonly MarkdownInlineStyle[];
      readonly paragraphId: string;
      readonly sourceStart: number;
      readonly sourceEnd: number;
      readonly exact: boolean;
    }
  | { readonly kind: 'boundary'; readonly value: MappedMarkdown };

interface MarkdownInlineContext {
  readonly tableCell: boolean;
  readonly hardBreakHtml?: boolean;
  /** A Markdown heading already communicates uniform bold emphasis. */
  readonly suppressBold?: boolean;
  readonly displayMode: SemanticLayout['displayMode'];
  readonly sourceScope: string;
  readonly sourceCapture?: MarkdownSourceCapture;
  /** Character style IDs whose runs project as code spans. */
  readonly codeCharacterStyles?: ReadonlySet<string>;
  readonly inlineControlRanges?: ReadonlyMap<
    string,
    readonly { readonly start: number; readonly end: number }[]
  >;
}

export interface MarkdownSourceCapture {
  /** Capture all text with enough sparse boundaries to resolve ordinary native positions. */
  readonly allSources?: boolean;
  /** Scopes whose cross-paragraph occurrences require every paragraph's source slices. */
  readonly allSourceScopes: ReadonlySet<string>;
  /** Sorted source boundaries needed for each reviewed paragraph. */
  readonly offsetsBySource: ReadonlyMap<string, readonly number[]>;
}

export function markdownSourceCaptureKey(sourceScope: string, paragraphId: string): string {
  return `${sourceScope}\0${paragraphId}`;
}

const MARKDOWN_PUNCTUATION = new Set('\\`*{}[]()#+-.!_|~=');
const MARKDOWN_UNICODE_PUNCTUATION = /[\p{P}\p{S}]/u;
const WORD_CHARACTER = /[\p{L}\p{N}]/u;

function isInWordHyphen(
  value: string,
  index: number,
  adjacent: { readonly before?: string; readonly after?: string }
): boolean {
  if (value[index] !== '-') return false;
  const before = index > 0 ? sourceEdge(value.slice(0, index), 'last') : adjacent.before;
  const after =
    index + 1 < value.length ? sourceEdge(value.slice(index + 1), 'first') : adjacent.after;
  return (
    before !== undefined &&
    after !== undefined &&
    WORD_CHARACTER.test(before) &&
    WORD_CHARACTER.test(after)
  );
}

function assertNever(value: never): never {
  throw new TypeError(`Unsupported Markdown inline style: ${String(value)}`);
}

/** Escape every file-derived character that can open Markdown or raw HTML. */
export function escapeText(value: string, tableCell = false): string {
  return escapedText(value, tableCell).markdown;
}

function escapedText(
  value: string,
  tableCell: boolean,
  requestedBoundaries: ReadonlySet<number> = new Set(),
  adjacent: { readonly before?: string; readonly after?: string } = {},
  complete = false,
  literal = false
): { readonly markdown: string; readonly boundaries: ReadonlyMap<number, number> } {
  if (literal) {
    // Code span content is copied verbatim, so every native offset is its own Markdown offset.
    const boundaries = new Map<number, number>();
    for (const offset of requestedBoundaries) {
      if (offset >= 0 && offset <= value.length) boundaries.set(offset, offset);
    }
    return { markdown: value, boundaries };
  }
  let markdown = '';
  const boundaries = new Map<number, number>();
  const record = (offset: number): void => {
    if (requestedBoundaries.has(offset)) boundaries.set(offset, markdown.length);
  };
  record(0);
  let index = 0;
  while (index < value.length) {
    const sourceStart = index;
    const markdownStart = markdown.length;
    const character = value[index]!;
    if (character === '\r' && value[index + 1] === '\n') {
      if (complete) boundaries.set(index, markdown.length);
      markdown += tableCell ? '<br>' : '  \n';
      index += 1;
      if (complete) boundaries.set(index, markdown.length);
      record(index);
      index += 1;
      if (complete) boundaries.set(index, markdown.length);
      record(index);
      continue;
    }
    if (character === '\r' || character === '\n' || character === '\f') {
      markdown += tableCell ? '<br>' : '  \n';
    } else if (character === '&') {
      markdown += '&amp;';
    } else if (character === '<') {
      markdown += '&lt;';
    } else if (character === '>') {
      markdown += '&gt;';
    } else if (character === '\t') {
      markdown += ' ';
    } else {
      if (MARKDOWN_PUNCTUATION.has(character) && !isInWordHyphen(value, index, adjacent)) {
        markdown += '\\';
      }
      markdown += character;
    }
    index += 1;
    if (complete && markdown.length - markdownStart !== index - sourceStart) {
      boundaries.set(sourceStart, markdownStart);
      boundaries.set(index, markdown.length);
    }
    record(index);
  }
  return { markdown, boundaries };
}

function mappedTextChunk(
  chunk: Extract<MarkdownInlineChunk, { kind: 'text' }>,
  tableCell: boolean,
  capture: MarkdownSourceCapture | undefined,
  sourceScope: string,
  adjacent: { readonly before?: string; readonly after?: string } = {}
): MappedMarkdown {
  const requested = capture?.offsetsBySource.get(
    markdownSourceCaptureKey(sourceScope, chunk.paragraphId)
  );
  if (
    !capture ||
    (!capture.allSources && !capture.allSourceScopes.has(sourceScope) && !requested)
  ) {
    return literalMarkdown(
      isCodeChunk(chunk)
        ? chunk.sourceText
        : escapedText(chunk.sourceText, tableCell, new Set(), adjacent).markdown
    );
  }
  const relativeBoundaries = new Set<number>([0, chunk.sourceText.length]);
  if (chunk.exact) {
    const offsets = requested ?? [];
    let low = 0;
    let high = offsets.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (offsets[middle]! < chunk.sourceStart) low = middle + 1;
      else high = middle;
    }
    for (let index = low; index < offsets.length; index += 1) {
      const offset = offsets[index]!;
      if (offset > chunk.sourceEnd) break;
      relativeBoundaries.add(offset - chunk.sourceStart);
    }
  }
  const escaped = escapedText(
    chunk.sourceText,
    tableCell,
    relativeBoundaries,
    adjacent,
    capture.allSources && chunk.exact,
    isCodeChunk(chunk)
  );
  const markdownBoundaries = chunk.exact
    ? [...escaped.boundaries]
        .sort(([left], [right]) => left - right)
        .map(([sourceOffset, markdownOffset]) => ({
          sourceOffset: chunk.sourceStart + sourceOffset,
          markdownOffset,
        }))
    : undefined;
  const source: MarkdownSourceSlice = {
    sourceScope,
    paragraphId: chunk.paragraphId,
    sourceStart: chunk.sourceStart,
    sourceEnd: chunk.sourceEnd,
    markdownStart: 0,
    markdownEnd: escaped.markdown.length,
    ...(markdownBoundaries ? { markdownBoundaries } : {}),
    exact: chunk.exact,
    ...(capture.allSources ? { completeBoundaries: true } : {}),
  };
  return { markdown: escaped.markdown, sources: [source] };
}

function markdownInlineStyles(
  span: StyleSpanRecord,
  context: MarkdownInlineContext
): readonly MarkdownInlineStyle[] {
  const deleted =
    context.displayMode === 'all-markup' &&
    span.revisions !== undefined &&
    revisionsAreDeletion(span.revisions);
  // Strike stays innermost. GFM refuses an intraword `~~` opener immediately before another
  // delimiter, while `*~~text~~*` and `**~~text~~**` remain valid in every transition direction.
  return [
    ...(span.style.italic ? (['italic'] as const) : []),
    ...(span.style.bold && !context.suppressBold ? (['bold'] as const) : []),
    ...(span.style.strike || span.style.doubleStrike || deleted ? (['strike'] as const) : []),
    ...(isCodeSpan(span, context) ? (['code'] as const) : []),
  ];
}

/** An explicit application code identity: the run's own `w:rStyle`, never its font. */
function isCodeSpan(span: StyleSpanRecord, context: MarkdownInlineContext): boolean {
  const styles = context.codeCharacterStyles;
  if (!styles || styles.size === 0 || span.equation !== undefined) return false;
  const reference = span.props.find((property) => property.localName === 'rStyle');
  const styleId = reference?.attributes?.val;
  return typeof styleId === 'string' && styles.has(styleId);
}

function isCodeChunk(chunk: MarkdownInlineChunk | undefined): boolean {
  return chunk?.kind === 'text' && chunk.styles[chunk.styles.length - 1] === 'code';
}

/** The first or last character a reader sees for a chunk; a code span shows its fence. */
function visibleEdge(chunk: MarkdownInlineChunk, edge: 'first' | 'last'): string | undefined {
  if (chunk.kind !== 'text') return sourceEdge(chunk.value.markdown, edge);
  return isCodeChunk(chunk) ? '`' : sourceEdge(chunk.sourceText, edge);
}

/** Styles that use flanking delimiters or tags; code spans have no flanking rules. */
function delimitedStyles(styles: readonly MarkdownInlineStyle[]): readonly MarkdownInlineStyle[] {
  return styles[styles.length - 1] === 'code' ? styles.slice(0, -1) : styles;
}

/**
 * CommonMark code span delimiters for one contiguous code group.
 *
 * The fence is the shortest backtick string whose length occurs nowhere inside the content.
 * One space pads both sides when the content touches a backtick, or when it begins and ends
 * with a space without being all spaces, because a parser strips exactly one such pair.
 */
function codeSpanDelimiters(content: string): { readonly open: string; readonly close: string } {
  const lengths = new Set<number>();
  for (const match of content.matchAll(/`+/gu)) lengths.add(match[0].length);
  let length = 1;
  while (lengths.has(length)) length += 1;
  const fence = '`'.repeat(length);
  const pad =
    content.startsWith('`') ||
    content.endsWith('`') ||
    (content.length > 1 &&
      content.startsWith(' ') &&
      content.endsWith(' ') &&
      /[^ ]/u.test(content))
      ? ' '
      : '';
  return { open: fence + pad, close: pad + fence };
}

/** Content of the code group starting at `index`: following chunks with identical styles. */
function codeGroupContent(
  chunks: readonly MarkdownInlineChunk[],
  index: number,
  end = chunks.length
): string {
  const first = chunks[index];
  if (first?.kind !== 'text') return '';
  let content = '';
  for (let at = index; at < end; at += 1) {
    const chunk = chunks[at];
    if (chunk?.kind !== 'text' || !sameInlineStyles(chunk.styles, first.styles)) break;
    content += chunk.sourceText;
  }
  return content;
}

function sameInlineStyles(
  left: readonly MarkdownInlineStyle[],
  right: readonly MarkdownInlineStyle[]
): boolean {
  return left.length === right.length && left.every((style, index) => style === right[index]);
}

function isBridgeableWhitespace(value: string): boolean {
  return value.length > 0 && /^[ \t]+$/u.test(value);
}

/**
 * Word commonly divides a visually continuous phrase into one run per word. `writeText` keeps
 * leading and trailing whitespace outside delimiters so an isolated run always produces valid
 * Markdown. When whitespace is surrounded by matching styled runs in the same paragraph, it is
 * safe—and materially cleaner—to retain that style across the run boundary.
 */
function bridgeMatchingStylesAcrossWhitespace(
  chunks: readonly MarkdownInlineChunk[]
): MarkdownInlineChunk[] {
  const bridged = [...chunks];
  let index = 0;
  while (index < chunks.length) {
    const chunk = chunks[index];
    if (
      chunk?.kind !== 'text' ||
      chunk.styles.length > 0 ||
      !isBridgeableWhitespace(chunk.sourceText)
    ) {
      index += 1;
      continue;
    }
    const start = index;
    let matchingWhitespaceStyles = true;
    while (index < chunks.length) {
      const candidate = chunks[index];
      if (
        candidate?.kind !== 'text' ||
        candidate.styles.length > 0 ||
        candidate.paragraphId !== chunk.paragraphId ||
        !isBridgeableWhitespace(candidate.sourceText)
      ) {
        break;
      }
      if (!sameInlineStyles(candidate.semanticStyles, chunk.semanticStyles)) {
        matchingWhitespaceStyles = false;
      }
      index += 1;
    }
    const previous = chunks[start - 1];
    const next = chunks[index];
    if (
      previous?.kind !== 'text' ||
      next?.kind !== 'text' ||
      previous.paragraphId !== chunk.paragraphId ||
      next.paragraphId !== chunk.paragraphId ||
      !matchingWhitespaceStyles ||
      delimitedStyles(previous.styles).length === 0 ||
      !sameInlineStyles(chunk.semanticStyles, delimitedStyles(previous.styles)) ||
      !sameInlineStyles(delimitedStyles(previous.styles), delimitedStyles(next.styles))
    ) {
      continue;
    }
    for (let whitespace = start; whitespace < index; whitespace += 1) {
      const candidate = chunks[whitespace];
      if (candidate?.kind === 'text')
        bridged[whitespace] = { ...candidate, styles: delimitedStyles(previous.styles) };
    }
  }
  return bridged;
}

function markdownStyleDelimiter(style: MarkdownInlineStyle): string {
  switch (style) {
    case 'strike':
      return '~~';
    case 'italic':
      return '*';
    case 'bold':
      return '**';
    case 'code':
      throw new TypeError('Code span delimiters depend on their content');
    case 'codeHtml':
      throw new TypeError('Structured code uses an HTML inline island');
    default:
      return assertNever(style);
  }
}

function markdownStyleTag(style: MarkdownInlineStyle): string {
  switch (style) {
    case 'strike':
      return 'del';
    case 'italic':
      return 'em';
    case 'bold':
      return 'strong';
    case 'code':
      throw new TypeError('Code spans keep backtick syntax inside HTML islands');
    case 'codeHtml':
      return 'code';
    default:
      return assertNever(style);
  }
}

function sharedStylePrefix(
  left: readonly MarkdownInlineStyle[],
  right: readonly MarkdownInlineStyle[]
): number {
  let retained = 0;
  while (retained < left.length && retained < right.length && left[retained] === right[retained]) {
    retained += 1;
  }
  return retained;
}

function sourceEdge(value: string, edge: 'first' | 'last'): string | undefined {
  if (value.length === 0) return undefined;
  if (edge === 'first') return String.fromCodePoint(value.codePointAt(0)!);
  const last = value.length - 1;
  const lastCodeUnit = value.charCodeAt(last);
  const previousCodeUnit = last > 0 ? value.charCodeAt(last - 1) : 0;
  const startsSurrogatePair =
    lastCodeUnit >= 0xdc00 &&
    lastCodeUnit <= 0xdfff &&
    previousCodeUnit >= 0xd800 &&
    previousCodeUnit <= 0xdbff;
  return value.slice(startsSurrogatePair ? last - 1 : last);
}

function chunkAdjacency(
  chunks: readonly MarkdownInlineChunk[],
  index: number
): { readonly before?: string; readonly after?: string } {
  const current = chunks[index];
  if (current?.kind !== 'text') return {};
  const previous = chunks[index - 1];
  const next = chunks[index + 1];
  return {
    ...(previous?.kind === 'text' && previous.paragraphId === current.paragraphId
      ? { before: sourceEdge(previous.sourceText, 'last') }
      : {}),
    ...(next?.kind === 'text' && next.paragraphId === current.paragraphId
      ? { after: sourceEdge(next.sourceText, 'first') }
      : {}),
  };
}

function inlineIslandNeedsHtml(
  chunks: readonly MarkdownInlineChunk[],
  start: number,
  end: number
): boolean {
  for (let index = start; index < end; index += 1) {
    const chunk = chunks[index]!;
    if (chunk.kind !== 'text') continue;
    if (chunk.styles.includes('codeHtml')) return true;
    const previous = chunks[index - 1];
    const next = chunks[index + 1];
    const styles = delimitedStyles(chunk.styles);
    const previousStyles =
      index > start && previous?.kind === 'text' ? delimitedStyles(previous.styles) : ([] as const);
    const nextStyles =
      index + 1 < end && next?.kind === 'text' ? delimitedStyles(next.styles) : ([] as const);
    const opensStyle = sharedStylePrefix(previousStyles, styles) < styles.length;
    const closesStyle = sharedStylePrefix(styles, nextStyles) < styles.length;
    const first = visibleEdge(chunk, 'first');
    const last = visibleEdge(chunk, 'last');
    const before = previous ? visibleEdge(previous, 'last') : undefined;
    const after = next ? visibleEdge(next, 'first') : undefined;
    // Backslash escaping punctuation does not change CommonMark delimiter flanking. Use fixed
    // semantic tags for this island when a delimiter would open or close against that boundary.
    if (
      opensStyle &&
      first &&
      before &&
      MARKDOWN_UNICODE_PUNCTUATION.test(first) &&
      !/\s/u.test(before)
    ) {
      return true;
    }
    if (
      closesStyle &&
      last &&
      after &&
      MARKDOWN_UNICODE_PUNCTUATION.test(last) &&
      !/\s/u.test(after)
    ) {
      return true;
    }
  }
  return false;
}

function htmlInlineIsland(
  chunks: readonly MarkdownInlineChunk[],
  start: number,
  end: number,
  tableCell: boolean,
  capture: MarkdownSourceCapture | undefined,
  sourceScope: string
): MappedMarkdown {
  const parts: MappedMarkdown[] = [];
  let openStyles: readonly MarkdownInlineStyle[] = [];
  let codeClose = '';
  const close = (style: MarkdownInlineStyle): string =>
    style === 'code' ? codeClose : `</${markdownStyleTag(style)}>`;
  for (let index = start; index < end; index += 1) {
    const chunk = chunks[index]!;
    if (chunk.kind !== 'text') continue;
    const retained = sharedStylePrefix(openStyles, chunk.styles);
    for (let styleIndex = openStyles.length - 1; styleIndex >= retained; styleIndex -= 1) {
      parts.push(literalMarkdown(close(openStyles[styleIndex]!)));
    }
    for (let styleIndex = retained; styleIndex < chunk.styles.length; styleIndex += 1) {
      const style = chunk.styles[styleIndex]!;
      if (style === 'code') {
        const delimiters = codeSpanDelimiters(codeGroupContent(chunks, index, end));
        codeClose = delimiters.close;
        parts.push(literalMarkdown(delimiters.open));
      } else parts.push(literalMarkdown(`<${markdownStyleTag(style)}>`));
    }
    parts.push(
      mappedTextChunk(chunk, tableCell, capture, sourceScope, chunkAdjacency(chunks, index))
    );
    openStyles = chunk.styles;
  }
  for (let index = openStyles.length - 1; index >= 0; index -= 1) {
    parts.push(literalMarkdown(close(openStyles[index]!)));
  }
  return concatMarkdown(parts);
}

/** Keep delimiter state across run, line, and page-fragment boundaries of one logical paragraph. */
export class MarkdownInlineWriter {
  readonly #context: MarkdownInlineContext;
  readonly #chunks: MarkdownInlineChunk[] = [];

  constructor(context: MarkdownInlineContext) {
    this.#context = context;
  }

  #writeSourceText(
    token: MarkdownTextToken,
    sourceText: string,
    relativeStart: number,
    styles: readonly MarkdownInlineStyle[],
    exact: boolean,
    semanticStyles: readonly MarkdownInlineStyle[] = styles
  ): void {
    if (sourceText.length === 0) return;
    this.#chunks.push({
      kind: 'text',
      sourceText,
      styles,
      semanticStyles,
      paragraphId: token.paragraphId,
      sourceStart: exact
        ? token.span.range.start + (token.sourceOffset ?? 0) + relativeStart
        : token.span.range.start,
      sourceEnd: exact
        ? token.span.range.start + (token.sourceOffset ?? 0) + relativeStart + sourceText.length
        : token.span.range.end,
      exact,
    });
  }

  writeText(token: MarkdownTextToken): void {
    // `trimStart`/`trimEnd` split on the same `\s` set as `/^(\s*)([\s\S]*?)(\s*)$/`, in
    // linear time. The lazy middle retried the trailing `\s*$` at every inner whitespace
    // run, which is quadratic on a long run that does not reach the end.
    const source = token.sourceText;
    const textStart = source.length - source.trimStart().length;
    const textEnd = textStart === source.length ? textStart : source.trimEnd().length;
    const leading = source.slice(0, textStart);
    const text = source.slice(textStart, textEnd);
    const trailing = source.slice(textEnd);
    const exact =
      token.exact ??
      (token.span.projected !== true &&
        token.span.equation === undefined &&
        token.sourceText.length === token.span.range.end - token.span.range.start);
    const nativeStyles = markdownInlineStyles(token.span, this.#context);
    // A content-control projection can replace only its native source span. Backticks
    // would turn that structural replacement into literal text. HTML keeps its ownership.
    const structured =
      nativeStyles.includes('code') &&
      this.#context.inlineControlRanges
        ?.get(token.paragraphId)
        ?.some((range) => token.span.range.start < range.end && token.span.range.end > range.start);
    const styles = structured
      ? nativeStyles.map((style) => (style === 'code' ? ('codeHtml' as const) : style))
      : nativeStyles;
    if (styles[styles.length - 1] === 'code') {
      // Code keeps its own whitespace. A manual break cannot sit inside a code span, so it
      // closes the span and keeps only the surrounding emphasis.
      const outer = styles.slice(0, -1);
      let offset = 0;
      for (const piece of source.split(/(\r\n|[\r\n\f])/u)) {
        this.#writeSourceText(
          token,
          piece,
          offset,
          /^[\r\n\f]/u.test(piece) ? outer : styles,
          exact
        );
        offset += piece.length;
      }
      return;
    }
    this.#writeSourceText(token, leading, 0, [], exact, styles);
    this.#writeSourceText(token, text, leading.length, styles, exact);
    this.#writeSourceText(token, trailing, leading.length + text.length, [], exact, styles);
  }

  writeBoundary(markdown: string): void {
    this.writeMappedBoundary(literalMarkdown(markdown));
  }

  writeMappedBoundary(value: MappedMarkdown): void {
    if (value.markdown.length === 0 && value.sources.length === 0) return;
    this.#chunks.push({ kind: 'boundary', value });
  }

  finishMapped(): MappedMarkdown {
    const chunks = bridgeMatchingStylesAcrossWhitespace(this.#chunks);
    const parts: MappedMarkdown[] = [];
    let openStyles: readonly MarkdownInlineStyle[] = [];
    let codeClose = '';
    const transition = (nextStyles: readonly MarkdownInlineStyle[], at = -1): void => {
      const retained = sharedStylePrefix(openStyles, nextStyles);
      for (let index = openStyles.length - 1; index >= retained; index -= 1) {
        const style = openStyles[index]!;
        parts.push(literalMarkdown(style === 'code' ? codeClose : markdownStyleDelimiter(style)));
      }
      for (let index = retained; index < nextStyles.length; index += 1) {
        const style = nextStyles[index]!;
        if (style === 'code') {
          const delimiters = codeSpanDelimiters(codeGroupContent(chunks, at));
          codeClose = delimiters.close;
          parts.push(literalMarkdown(delimiters.open));
        } else parts.push(literalMarkdown(markdownStyleDelimiter(style)));
      }
      openStyles = nextStyles;
    };
    for (let index = 0; index < chunks.length; index += 1) {
      const chunk = chunks[index]!;
      if (chunk.kind === 'boundary') {
        transition([]);
        parts.push(chunk.value);
        continue;
      }
      if (chunk.styles.length === 0) {
        transition([]);
        parts.push(this.#mappedSourceChunk(chunk, index));
        continue;
      }
      let end = index + 1;
      while (end < chunks.length) {
        const candidate = chunks[end];
        if (candidate?.kind !== 'text' || candidate.styles.length === 0) break;
        end += 1;
      }
      if (inlineIslandNeedsHtml(chunks, index, end)) {
        transition([]);
        parts.push(
          htmlInlineIsland(
            chunks,
            index,
            end,
            this.#context.tableCell || this.#context.hardBreakHtml === true,
            this.#context.sourceCapture,
            this.#context.sourceScope
          )
        );
        index = end - 1;
        continue;
      }
      // The island was checked as a whole, with real neighbouring styles. Re-checking a suffix
      // would treat a retained style as newly opened at a code fence or punctuation.
      for (let at = index; at < end; at += 1) {
        const member = chunks[at] as Extract<MarkdownInlineChunk, { kind: 'text' }>;
        transition(member.styles, at);
        parts.push(this.#mappedSourceChunk(member, at, chunks));
      }
      index = end - 1;
    }
    transition([]);
    return concatMarkdown(parts);
  }

  #mappedSourceChunk(
    chunk: Extract<MarkdownInlineChunk, { kind: 'text' }>,
    index: number,
    chunks: readonly MarkdownInlineChunk[] = this.#chunks
  ): MappedMarkdown {
    const tableCell = this.#context.tableCell || this.#context.hardBreakHtml === true;
    return mappedTextChunk(
      chunk,
      tableCell,
      this.#context.sourceCapture,
      this.#context.sourceScope,
      chunkAdjacency(chunks, index)
    );
  }

  finish(): string {
    return this.finishMapped().markdown;
  }
}
