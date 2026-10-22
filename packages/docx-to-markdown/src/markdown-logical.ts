import type { MediaRendering, AnchorProjection } from './markdown-media.ts';
// Logical-block grouping and table emission over published semantic records.
//
// Split-table fragments merge back into logical rows here, and both GFM and nested-table
// serialization live beside that merge. Rendering a CELL's content recurses into the full
// block renderer, which lives in `markdown.ts`; callers thread it in as `render` so the
// dependency stays one-directional (markdown.ts imports this module, never the reverse).

import type {
  BlockFragmentRecord,
  ParagraphFragmentRecord,
  SemanticLayout,
  TableFragmentRecord,
} from '@docx-editor.dev/core/layout';
import {
  concatMarkdown,
  EMPTY_MAPPED_MARKDOWN,
  escapeUnescapedTablePipes,
  literalMarkdown,
  replaceNewlinesWithHtmlBreaks,
  type MappedMarkdown,
} from './markdown-source-map.ts';
import { nestedTableHtml, tableWidth } from './markdown-nested-table.ts';
import type { MarkdownSourceCapture } from './markdown-inline.ts';

export interface TranslationContext {
  readonly media?: MediaRendering;
  readonly anchors?: AnchorProjection;
  readonly noteLabelByScope: Map<string, string>;
  readonly tableCell: boolean;
  readonly tableSyntax?: 'gfm' | 'html' | 'auto';
  readonly codeCharacterStyles?: ReadonlySet<string>;
  readonly hardBreakHtml?: boolean;
  readonly suppressBold?: boolean;
  readonly displayMode: SemanticLayout['displayMode'];
  /** Translator-local story/part scope disambiguating paragraph ids across DOCX parts. */
  readonly sourceScope: string;
  readonly listIndentByParagraphId: ReadonlyMap<string, number>;
  readonly listMarkerByParagraphId: ReadonlyMap<
    string,
    NonNullable<ParagraphFragmentRecord['marker']>
  >;
  readonly tablesById: ReadonlyMap<string, TableProjection>;
  readonly sourceCapture?: MarkdownSourceCapture;
  /** Labels emitted into the current page body, when page-local note visibility is tracked. */
  readonly emittedNoteLabels?: Set<string>;
  readonly pageIndex?: number;
}

export interface LogicalParagraph {
  readonly kind: 'paragraph';
  readonly fragments: ParagraphFragmentRecord[];
}

export interface LogicalTable {
  readonly kind: 'table';
  readonly fragments: TableFragmentRecord[];
}

export type LogicalBlock = LogicalParagraph | LogicalTable;

export interface LogicalCell {
  readonly blocks: BlockFragmentRecord[];
  readonly gridSpan: number;
  readonly vMergeContinue: boolean;
  readonly gridColumn: number;
}

export interface LogicalRow {
  readonly id: string;
  readonly isHeaderRow: boolean;
  readonly isHeaderRepeat: boolean;
  readonly cells: LogicalCell[];
}

export interface TableProjection {
  readonly fragments: TableFragmentRecord[];
  readonly columnCount: number;
}

/** The full block renderer from `markdown.ts`, threaded in to render cell content. */
export type RenderLogicalBlocks = (
  blocks: readonly LogicalBlock[],
  context: TranslationContext,
  nested?: boolean,
  pageScoped?: boolean
) => MappedMarkdown;

export function assertNever(value: never): never {
  throw new TypeError(`Unsupported semantic record: ${JSON.stringify(value)}`);
}

export function logicalBlocks(blocks: readonly BlockFragmentRecord[]): LogicalBlock[] {
  const result: LogicalBlock[] = [];
  for (const block of blocks) {
    const previous = result[result.length - 1];
    switch (block.kind) {
      case 'paragraph':
        if (
          previous?.kind === 'paragraph' &&
          previous.fragments[0]?.paragraphId === block.paragraphId
        ) {
          previous.fragments.push(block);
        } else {
          result.push({ kind: 'paragraph', fragments: [block] });
        }
        break;
      case 'table':
        if (previous?.kind === 'table' && previous.fragments[0]?.tableId === block.tableId) {
          previous.fragments.push(block);
        } else {
          result.push({ kind: 'table', fragments: [block] });
        }
        break;
      default:
        assertNever(block);
    }
  }
  return result;
}

export function mergeRows(
  fragments: readonly TableFragmentRecord[],
  includeHeaderRepeats: boolean
): LogicalRow[] {
  const rows: LogicalRow[] = [];
  const byId = new Map<string, LogicalRow>();
  for (const fragment of fragments) {
    for (const row of fragment.rows) {
      if (row.isHeaderRepeat && !includeHeaderRepeats) continue;
      let logical = byId.get(row.id);
      // A tall table can restart its repeating header inside another fragment on the same physical
      // page. That is a visual occurrence of the row already collected, not additional cell
      // content. Row continuations remain mergeable because they are not header repeats.
      if (logical && includeHeaderRepeats && row.isHeaderRepeat) continue;
      if (!logical) {
        logical = {
          id: row.id,
          isHeaderRow: row.isHeaderRow,
          isHeaderRepeat: row.isHeaderRepeat,
          cells: [],
        };
        byId.set(row.id, logical);
        rows.push(logical);
      }
      for (const [index, cell] of row.cells.entries()) {
        const gridColumn =
          cell.resolvedGridColumn ?? cell.logicalGridColumn ?? cell.gridColumn ?? index;
        const existing = logical.cells.find((candidate) => candidate.gridColumn === gridColumn);
        if (existing) existing.blocks.push(...cell.blocks);
        else {
          logical.cells.push({
            blocks: [...cell.blocks],
            gridSpan: cell.resolvedGridSpan ?? cell.gridSpan,
            vMergeContinue: cell.vMergeContinue,
            gridColumn,
          });
        }
      }
      logical.cells.sort((left, right) => left.gridColumn - right.gridColumn);
    }
  }
  return rows;
}

function cellAlignment(cell: LogicalCell): ParagraphFragmentRecord['alignment'] {
  for (const block of cell.blocks) {
    if (block.kind === 'paragraph') return block.alignment;
  }
  return 'left';
}

/**
 * Column count for emission. The declared grid is preferred, but a row can carry more cells
 * than `w:tblGrid` declares; widening to the widest laid-out cell keeps every painted cell in
 * the output instead of silently dropping the overflow.
 */
function emittedTableWidth(
  rows: readonly LogicalRow[],
  projection: TableProjection | undefined,
  fragments: readonly TableFragmentRecord[]
): number {
  let resolvedWidth: number | undefined;
  for (const fragment of projection?.fragments ?? fragments)
    for (const row of fragment.rows)
      for (const cell of row.cells)
        if (cell.resolvedGridColumnCount !== undefined)
          resolvedWidth = cell.resolvedGridColumnCount;
  let width =
    resolvedWidth ?? tableWidth(projection?.columnCount, fragments[0]?.columnEdges.length);
  for (const row of rows) {
    for (const cell of row.cells) {
      width = Math.max(width, cell.gridColumn + Math.max(cell.gridSpan, 1));
    }
  }
  return width;
}

function withTableSource(
  value: MappedMarkdown,
  tableId: string | undefined,
  context: TranslationContext
): MappedMarkdown {
  if (!tableId || !context.sourceCapture?.allSources) return value;
  return {
    ...value,
    tables: [
      {
        sourceScope: context.sourceScope,
        tableId,
        markdownStart: 0,
        markdownEnd: value.markdown.length,
      },
      ...(value.tables ?? []),
    ],
  };
}

/** Auto syntax preserves structures that a GFM cell cannot express. */
export function tableNeedsHtml(
  fragments: readonly TableFragmentRecord[],
  context: TranslationContext
): boolean {
  const id = fragments[0]?.tableId,
    complete = id ? context.tablesById.get(id)?.fragments : undefined;
  return mergeRows(complete ?? fragments, false).some((row) =>
    row.cells.some(
      (cell) =>
        cell.gridSpan > 1 ||
        cell.vMergeContinue ||
        logicalBlocks(cell.blocks).length > 1 ||
        cell.blocks.some((block) => block.kind === 'table')
    )
  );
}

/** GFM's unsupported nested-table shape uses mapped inline HTML. */
export function nestedTableMarkdown(
  fragments: readonly TableFragmentRecord[],
  context: TranslationContext,
  pageScoped: boolean,
  render: RenderLogicalBlocks
): MappedMarkdown {
  const tableId = fragments[0]?.tableId;
  const projection = tableId ? context.tablesById.get(tableId) : undefined;
  const completeFragments = projection?.fragments ?? fragments;
  const rows = pageScoped ? mergeRows(fragments, true) : mergeRows(completeFragments, false);
  if (rows.length === 0) return EMPTY_MAPPED_MARKDOWN;
  const width = emittedTableWidth(rows, projection, fragments);
  const nestedContext = { ...context, tableCell: true };
  const rowIndexes = new Map(rows.map((row, index) => [row, index]));
  const rowCells = new Map(
    rows.map((row) => [row, new Map(row.cells.map((cell) => [cell.gridColumn, cell]))])
  );
  const markdown = nestedTableHtml(
    rows,
    width,
    (row, columnIndex) => {
      const cell = rowCells.get(row)?.get(columnIndex);
      return !cell || cell.vMergeContinue
        ? EMPTY_MAPPED_MARKDOWN
        : render(logicalBlocks(cell.blocks), nestedContext, true, pageScoped);
    },
    context.tableSyntax !== undefined && context.tableSyntax !== 'gfm',
    context.tableSyntax === undefined || context.tableSyntax === 'gfm'
      ? undefined
      : (row, columnIndex) => {
          const cell = rowCells.get(row)?.get(columnIndex);
          if (!cell)
            return {
              columnSpan: 1,
              rowSpan: 1,
              skip: row.cells.some(
                (candidate) =>
                  candidate.gridColumn < columnIndex &&
                  candidate.gridColumn + candidate.gridSpan > columnIndex
              ),
            };
          if (cell.vMergeContinue) return { columnSpan: cell.gridSpan, rowSpan: 1, skip: true };
          let rowSpan = 1;
          const index = rowIndexes.get(row)!;
          while (
            index + rowSpan < rows.length &&
            rowCells.get(rows[index + rowSpan]!)?.get(columnIndex)?.vMergeContinue
          )
            rowSpan++;
          return { columnSpan: cell.gridSpan, rowSpan, skip: cell.vMergeContinue };
        }
  );
  return withTableSource(markdown, tableId, context);
}

function cellValues(
  row: LogicalRow,
  context: TranslationContext,
  pageScoped: boolean,
  render: RenderLogicalBlocks
): MappedMarkdown[] {
  const values: MappedMarkdown[] = [];
  const nestedContext = { ...context, tableCell: true };
  for (const cell of row.cells) {
    const value = cell.vMergeContinue
      ? EMPTY_MAPPED_MARKDOWN
      : replaceNewlinesWithHtmlBreaks(
          escapeUnescapedTablePipes(
            render(logicalBlocks(cell.blocks), nestedContext, true, pageScoped)
          )
        );
    while (values.length < cell.gridColumn) values.push(EMPTY_MAPPED_MARKDOWN);
    values[cell.gridColumn] = value;
    for (let span = 1; span < cell.gridSpan; span += 1) {
      values[cell.gridColumn + span] = EMPTY_MAPPED_MARKDOWN;
    }
  }
  return values;
}

export function tableMarkdown(
  fragments: readonly TableFragmentRecord[],
  context: TranslationContext,
  pageScoped: boolean,
  render: RenderLogicalBlocks
): MappedMarkdown {
  const tableId = fragments[0]?.tableId;
  const projection = tableId ? context.tablesById.get(tableId) : undefined;
  const completeFragments = projection?.fragments ?? fragments;
  const rows = pageScoped ? mergeRows(fragments, true) : mergeRows(completeFragments, false);
  if (rows.length === 0) return EMPTY_MAPPED_MARKDOWN;
  const width = emittedTableWidth(rows, projection, fragments);
  const normalize = (
    values: MappedMarkdown[],
    fallback: MappedMarkdown = EMPTY_MAPPED_MARKDOWN
  ): MappedMarkdown[] => Array.from({ length: width }, (_, index) => values[index] ?? fallback);
  const line = (row: LogicalRow): MappedMarkdown =>
    concatMarkdown(
      [
        literalMarkdown('| '),
        concatMarkdown(normalize(cellValues(row, context, pageScoped, render)), ' | '),
        literalMarkdown(' |'),
      ],
      ''
    );
  // GFM requires the first emitted row to be its header row. Preserve authored order: OOXML's
  // tblHeader flag is only an effective repeated header on the contiguous table prefix, and a
  // malformed/later flag must never pull that row ahead of preceding data.
  const header = rows[0]!;
  const alignments: string[] = [];
  for (const cell of header.cells) {
    const token = cellAlignment(cell);
    const rule = token === 'center' ? ':---:' : token === 'right' ? '---:' : '---';
    while (alignments.length < cell.gridColumn) alignments.push('---');
    alignments[cell.gridColumn] = rule;
    for (let span = 1; span < cell.gridSpan; span += 1) {
      alignments[cell.gridColumn + span] = '---';
    }
  }
  const markdown = concatMarkdown(
    [
      line(header),
      literalMarkdown(
        `| ${Array.from({ length: width }, (_, index) => alignments[index] ?? '---').join(' | ')} |`
      ),
      ...rows.slice(1).map(line),
    ],
    '\n'
  );
  return withTableSource(markdown, tableId, context);
}

export function indexTableBlocks(
  blocks: readonly BlockFragmentRecord[],
  tables: Map<string, TableProjection>
): void {
  for (const block of blocks) {
    if (block.kind === 'paragraph') continue;
    let projection = tables.get(block.tableId);
    if (!projection) {
      projection = {
        fragments: [],
        columnCount: Math.max(1, block.columnEdges.length - 1),
      };
      tables.set(block.tableId, projection);
    }
    projection.fragments.push(block);
    for (const row of block.rows) {
      if (row.isHeaderRepeat) continue;
      for (const cell of row.cells) indexTableBlocks(cell.blocks, tables);
    }
  }
}
