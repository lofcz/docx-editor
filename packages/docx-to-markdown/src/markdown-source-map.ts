import type {
  MarkdownSourceBoundary,
  MarkdownSourceMap,
  MarkdownSourceParagraph,
  MarkdownSourceTable,
  MarkdownSourceSlice as PublicMarkdownSourceSlice,
} from './markdown-source-types.ts';
export type { MarkdownSourceParagraph } from './markdown-source-types.ts';

/** Internal slices can track either all boundaries or only requested review boundaries. */
export interface MarkdownSourceSlice extends PublicMarkdownSourceSlice {
  readonly completeBoundaries?: boolean;
}

/** Markdown plus the source-backed substrings needed to bind review artifacts. */
export interface MappedMarkdown {
  readonly markdown: string;
  readonly sources: readonly MarkdownSourceSlice[];
  readonly paragraphs?: readonly MarkdownSourceParagraph[];
  readonly tables?: readonly MarkdownSourceTable[];
}

export const EMPTY_MAPPED_MARKDOWN: MappedMarkdown = Object.freeze({
  markdown: '',
  sources: Object.freeze([]),
});

export function freezeMarkdownSourceMap(value: MappedMarkdown): MarkdownSourceMap {
  const compact: PublicMarkdownSourceSlice[] = [];
  let current: PublicMarkdownSourceSlice | undefined;
  let boundaries: MarkdownSourceBoundary[] = [];
  const finish = (): void => {
    if (!current) return;
    compact.push(
      Object.freeze({
        ...current,
        ...(current.exact
          ? {
              markdownBoundaries: Object.freeze(
                boundaries.map((boundary) => Object.freeze(boundary))
              ),
            }
          : {}),
      })
    );
  };
  for (const { completeBoundaries: _complete, ...source } of value.sources) {
    if (
      current?.exact &&
      source.exact &&
      current.sourceScope === source.sourceScope &&
      current.paragraphId === source.paragraphId &&
      current.sourceEnd === source.sourceStart &&
      current.markdownEnd === source.markdownStart
    ) {
      current = { ...current, sourceEnd: source.sourceEnd, markdownEnd: source.markdownEnd };
    } else {
      finish();
      current = source;
      boundaries = [];
    }
    for (const boundary of source.markdownBoundaries ?? []) {
      const previous = boundaries[boundaries.length - 1];
      if (
        previous?.sourceOffset === boundary.sourceOffset &&
        previous.markdownOffset === boundary.markdownOffset
      )
        continue;
      const before = boundaries[boundaries.length - 2];
      if (
        before &&
        previous &&
        previous.sourceOffset - before.sourceOffset ===
          previous.markdownOffset - before.markdownOffset &&
        boundary.sourceOffset - previous.sourceOffset ===
          boundary.markdownOffset - previous.markdownOffset
      ) {
        boundaries[boundaries.length - 1] = { ...boundary };
      } else boundaries.push({ ...boundary });
    }
  }
  finish();
  return Object.freeze({
    unit: 'utf16-code-unit',
    sources: Object.freeze(compact),
    paragraphs: Object.freeze(
      (value.paragraphs ?? []).map((paragraph) => Object.freeze({ ...paragraph }))
    ),
    tables: Object.freeze((value.tables ?? []).map((table) => Object.freeze({ ...table }))),
  });
}

function transformedBoundaries(
  source: MarkdownSourceSlice,
  boundaryMap: readonly number[],
  changed: readonly number[]
): readonly MarkdownSourceBoundary[] | undefined {
  const boundaries = source.markdownBoundaries;
  if (!boundaries) return undefined;
  if (!source.completeBoundaries)
    return boundaries.map((boundary) => ({
      ...boundary,
      markdownOffset: boundaryMap[boundary.markdownOffset]!,
    }));
  // A transform can insert syntax inside a formerly linear text interval. Retain boundaries
  // around each insertion so downstream callers never interpolate through generated syntax.
  const expanded: MarkdownSourceBoundary[] = [];
  let changedIndex = 0;
  let high = changed.length;
  while (changedIndex < high) {
    const middle = (changedIndex + high) >>> 1;
    if (changed[middle]! < source.markdownStart) changedIndex = middle + 1;
    else high = middle;
  }
  for (const [index, boundary] of boundaries.entries()) {
    const previous = boundaries[index - 1];
    if (
      previous &&
      boundary.markdownOffset - previous.markdownOffset ===
        boundary.sourceOffset - previous.sourceOffset
    ) {
      while (changedIndex < changed.length && changed[changedIndex]! < boundary.markdownOffset) {
        const offset = changed[changedIndex++]!;
        if (offset > previous.markdownOffset)
          expanded.push({
            sourceOffset: previous.sourceOffset + offset - previous.markdownOffset,
            markdownOffset: boundaryMap[offset]!,
          });
      }
    }
    expanded.push({ ...boundary, markdownOffset: boundaryMap[boundary.markdownOffset]! });
    while (changedIndex < changed.length && changed[changedIndex]! <= boundary.markdownOffset)
      changedIndex += 1;
  }
  return expanded;
}

export function literalMarkdown(markdown: string): MappedMarkdown {
  return markdown.length === 0 ? EMPTY_MAPPED_MARKDOWN : { markdown, sources: [] };
}

function shiftedSources(
  sources: readonly MarkdownSourceSlice[],
  offset: number
): MarkdownSourceSlice[] {
  if (offset === 0) return [...sources];
  return sources.map((source) => ({
    ...source,
    markdownStart: source.markdownStart + offset,
    markdownEnd: source.markdownEnd + offset,
    markdownBoundaries: source.markdownBoundaries?.map((boundary) => ({
      ...boundary,
      markdownOffset: boundary.markdownOffset + offset,
    })),
  }));
}

export function shiftedParagraphs(
  paragraphs: readonly MarkdownSourceParagraph[],
  offset: number
): MarkdownSourceParagraph[] {
  return paragraphs.map((paragraph) => ({
    ...paragraph,
    ...(paragraph.markdownStart === undefined
      ? {}
      : { markdownStart: paragraph.markdownStart + offset }),
    ...(paragraph.markdownEnd === undefined ? {} : { markdownEnd: paragraph.markdownEnd + offset }),
  }));
}

function shiftedTables(
  tables: readonly MarkdownSourceTable[],
  offset: number
): MarkdownSourceTable[] {
  return tables.map((table) => ({
    ...table,
    markdownStart: table.markdownStart + offset,
    markdownEnd: table.markdownEnd + offset,
  }));
}

export function concatMarkdown(values: readonly MappedMarkdown[], separator = ''): MappedMarkdown {
  let markdown = '';
  const sources: MarkdownSourceSlice[] = [];
  const paragraphs: MarkdownSourceParagraph[] = [];
  const tables: MarkdownSourceTable[] = [];
  for (const [index, value] of values.entries()) {
    if (index > 0) markdown += separator;
    for (const source of shiftedSources(value.sources, markdown.length)) sources.push(source);
    paragraphs.push(...shiftedParagraphs(value.paragraphs ?? [], markdown.length));
    tables.push(...shiftedTables(value.tables ?? [], markdown.length));
    markdown += value.markdown;
  }
  return {
    markdown,
    sources,
    ...(paragraphs.length > 0 ? { paragraphs } : {}),
    ...(tables.length > 0 ? { tables } : {}),
  };
}

export function wrapMarkdown(value: MappedMarkdown, prefix: string, suffix = ''): MappedMarkdown {
  return {
    markdown: prefix + value.markdown + suffix,
    sources: shiftedSources(value.sources, prefix.length),
    ...(value.paragraphs ? { paragraphs: shiftedParagraphs(value.paragraphs, prefix.length) } : {}),
    ...(value.tables ? { tables: shiftedTables(value.tables, prefix.length) } : {}),
  };
}

/**
 * Apply a text transformation while retaining a map for every boundary in the input string.
 * `boundaryMap[n]` is the output offset immediately after consuming `n` input code units.
 */
export function transformMarkdown(
  value: MappedMarkdown,
  transform: (input: string) => {
    readonly markdown: string;
    readonly boundaryMap: readonly number[];
  }
): MappedMarkdown {
  const transformed = transform(value.markdown);
  if (transformed.boundaryMap.length !== value.markdown.length + 1) {
    throw new TypeError('Markdown transform returned an invalid boundary map');
  }
  const changed: number[] = [];
  if (value.sources.some((source) => source.completeBoundaries)) {
    for (let index = 1; index < transformed.boundaryMap.length; index += 1) {
      if (transformed.boundaryMap[index]! - transformed.boundaryMap[index - 1]! !== 1) {
        if (changed[changed.length - 1] !== index - 1) changed.push(index - 1);
        changed.push(index);
      }
    }
  }
  return {
    markdown: transformed.markdown,
    ...(value.tables
      ? {
          tables: value.tables.map((table) => ({
            ...table,
            markdownStart: transformed.boundaryMap[table.markdownStart]!,
            markdownEnd: transformed.boundaryMap[table.markdownEnd]!,
          })),
        }
      : {}),
    sources: value.sources.map((source) => ({
      ...source,
      markdownStart: transformed.boundaryMap[source.markdownStart]!,
      markdownEnd: transformed.boundaryMap[source.markdownEnd]!,
      markdownBoundaries: transformedBoundaries(source, transformed.boundaryMap, changed),
    })),
    ...(value.paragraphs
      ? {
          paragraphs: value.paragraphs.map((paragraph) => ({
            ...paragraph,
            ...(paragraph.markdownStart === undefined
              ? {}
              : { markdownStart: transformed.boundaryMap[paragraph.markdownStart]! }),
            ...(paragraph.markdownEnd === undefined
              ? {}
              : { markdownEnd: transformed.boundaryMap[paragraph.markdownEnd]! }),
          })),
        }
      : {}),
  };
}

export function withSourceParagraphs(
  value: MappedMarkdown,
  paragraphs: readonly MarkdownSourceParagraph[]
): MappedMarkdown {
  if (paragraphs.length === 0) return value;
  return { ...value, paragraphs: Object.freeze([...paragraphs]) };
}

export function preserveLeadingWhitespace(value: MappedMarkdown): MappedMarkdown {
  const leadingLength = /^[ \t]+/.exec(value.markdown)?.[0].length ?? 0;
  if (leadingLength === 0) return value;
  return transformMarkdown(value, (input) => {
    let markdown = '';
    const boundaryMap = [0];
    for (let index = 0; index < input.length; index += 1) {
      markdown += index < leadingLength ? '\u00a0' : input[index]!;
      boundaryMap.push(markdown.length);
    }
    return { markdown, boundaryMap };
  });
}

export function escapeUnescapedTablePipes(value: MappedMarkdown): MappedMarkdown {
  return transformMarkdown(value, (input) => {
    let markdown = '';
    let precedingBackslashes = 0;
    const boundaryMap = [0];
    for (let index = 0; index < input.length; index += 1) {
      const character = input[index]!;
      if (character === '\\') {
        markdown += character;
        precedingBackslashes += 1;
      } else {
        if (character === '|' && precedingBackslashes % 2 === 0) markdown += '\\';
        markdown += character;
        precedingBackslashes = 0;
      }
      boundaryMap[index + 1] = markdown.length;
    }
    return { markdown, boundaryMap };
  });
}

export function replaceNewlinesWithHtmlBreaks(value: MappedMarkdown): MappedMarkdown {
  return transformMarkdown(value, (input) => {
    let markdown = '';
    const boundaryMap = [0];
    let index = 0;
    while (index < input.length) {
      if (input[index] !== '\n') {
        markdown += input[index]!;
        index += 1;
        boundaryMap[index] = markdown.length;
        continue;
      }
      const start = index;
      while (index < input.length && input[index] === '\n') index += 1;
      markdown += '<br>';
      for (let consumed = start + 1; consumed <= index; consumed += 1) {
        boundaryMap[consumed] = markdown.length;
      }
    }
    return { markdown, boundaryMap };
  });
}

export function indentContinuationLines(
  value: MappedMarkdown,
  indentation: string
): MappedMarkdown {
  return transformMarkdown(value, (input) => {
    let markdown = '';
    const boundaryMap = [0];
    for (let index = 0; index < input.length; index += 1) {
      const character = input[index]!;
      markdown += character;
      if (character === '\n') markdown += indentation;
      boundaryMap.push(markdown.length);
    }
    return { markdown, boundaryMap };
  });
}

export function quoteMarkdownLines(value: MappedMarkdown): MappedMarkdown {
  return transformMarkdown(value, (input) => {
    const prefixAt = (index: number): string =>
      index === input.length || input[index] === '\n' ? '>' : '> ';
    let markdown = prefixAt(0);
    const boundaryMap = [markdown.length];
    for (let index = 0; index < input.length; index += 1) {
      const character = input[index]!;
      markdown += character;
      if (character === '\n') markdown += prefixAt(index + 1);
      boundaryMap.push(markdown.length);
    }
    return { markdown, boundaryMap };
  });
}
