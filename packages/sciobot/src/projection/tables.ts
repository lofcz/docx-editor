import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import type { BlockFragmentRecord } from '@docx-editor.dev/core/layout';
import type { MarkdownSourceSlice } from '@docx-editor.dev/docx-to-markdown';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { markdownPointAt, projectMarkdown, type MarkdownProjection } from './markdown';

export interface ProjectedTable {
  readonly id: string;
  readonly parentCellId: string | null;
  readonly rowIds: readonly string[];
}

export interface ProjectedTableCell {
  readonly id: string;
  readonly tableId: string;
  readonly rowId: string;
  readonly rowIndex: number;
  readonly gridColumn: number;
  readonly gridSpan: number;
  readonly verticalContinuation: boolean;
  readonly paragraphIds: readonly string[];
  /** Source-backed extents, with native escaping and fallback precision retained. */
  readonly sources: readonly MarkdownSourceSlice[];
}

export interface TableBoundary {
  readonly tableId: string;
  readonly role: 'cell' | 'row';
  readonly beforeCellId: string;
  readonly afterCellId: string;
  /** Empty or hidden cells have no text extent. Their structural identities remain present. */
  readonly extent: { readonly start: number; readonly end: number } | null;
}

export interface TableProjection extends MarkdownProjection {
  readonly tables: readonly ProjectedTable[];
  readonly cells: readonly ProjectedTableCell[];
  readonly boundaries: readonly TableBoundary[];
}

interface MutableCell extends Omit<ProjectedTableCell, 'paragraphIds' | 'sources'> {
  paragraphIds: Set<string>;
  sources: MarkdownSourceSlice[];
}

/** Use native resolved cell records. Do not reimplement Word grids, merges, or pagination. */
export function projectTables(
  layout: ExportSemanticLayout,
  revision: RevisionToken
): { readonly success: true; readonly projection: TableProjection } | NativeEditFailure {
  return projectTableOwnership(layout, projectMarkdown(layout, revision));
}

/** Attach native table identities to an already composed Markdown view without exporting it again. */
export function projectTableOwnership(
  layout: ExportSemanticLayout,
  projection: MarkdownProjection
): { readonly success: true; readonly projection: TableProjection } | NativeEditFailure {
  const tables = new Map<
    string,
    { id: string; parentCellId: string | null; rowIds: Set<string> }
  >();
  const cells = new Map<string, MutableCell>();
  const cellsByTable = new Map<string, MutableCell[]>();
  const paragraphCells = new Map<string, string>();
  let visits = 0;
  const visit = (
    blocks: readonly BlockFragmentRecord[],
    parentCellId: string | null,
    depth: number
  ): boolean => {
    if (depth > 64) return false;
    for (const block of blocks) {
      if (++visits > 100_000) return false;
      if (block.kind === 'paragraph') {
        if (!parentCellId) continue;
        const cell = cells.get(parentCellId)!;
        const ids = new Set([block.paragraphId]);
        for (const line of block.lines)
          for (const span of line.spans) ids.add(span.range.paragraphId);
        for (const id of ids) {
          const previous = paragraphCells.get(id);
          if (previous && previous !== parentCellId) return false;
          paragraphCells.set(id, parentCellId);
          cell.paragraphIds.add(id);
        }
        continue;
      }
      let table = tables.get(block.tableId);
      if (!table) {
        table = { id: block.tableId, parentCellId, rowIds: new Set() };
        tables.set(block.tableId, table);
      } else if (table.parentCellId !== parentCellId) return false;
      for (const row of block.rows) {
        if (++visits > 100_000) return false;
        table.rowIds.add(row.id);
        for (const record of row.cells) {
          if (++visits > 100_000) return false;
          let cell = cells.get(record.id);
          if (!cell) {
            cell = {
              id: record.id,
              tableId: block.tableId,
              rowId: row.id,
              rowIndex: row.rowIndex,
              gridColumn: record.resolvedGridColumn ?? record.gridColumn,
              gridSpan: record.resolvedGridSpan ?? record.gridSpan,
              verticalContinuation: record.vMergeContinue,
              paragraphIds: new Set(),
              sources: [],
            };
            cells.set(record.id, cell);
            const siblings = cellsByTable.get(block.tableId) ?? [];
            siblings.push(cell);
            cellsByTable.set(block.tableId, siblings);
          } else if (cell.tableId !== block.tableId || cell.rowId !== row.id) return false;
          if (!visit(record.blocks, cell.id, depth + 1)) return false;
        }
      }
    }
    return true;
  };
  for (const page of layout.pages) {
    if (!visit(page.fragments, null, 0))
      return {
        success: false,
        code: 'invalid-document',
        error: 'Table source records are ambiguous or exceed traversal limits.',
      };
  }
  for (const source of projection.sourceMap.sources) {
    if (source.sourceScope !== 'body') continue;
    const cellId = paragraphCells.get(source.paragraphId);
    if (cellId) cells.get(cellId)!.sources.push(source);
  }
  const boundaries: TableBoundary[] = [];
  for (const table of tables.values()) {
    const ordered = [...(cellsByTable.get(table.id) ?? [])].sort(
      (left, right) => left.rowIndex - right.rowIndex || left.gridColumn - right.gridColumn
    );
    for (let index = 1; index < ordered.length; index += 1) {
      const before = ordered[index - 1]!;
      const after = ordered[index]!;
      const start = before.sources.at(-1)?.markdownEnd;
      const end = after.sources[0]?.markdownStart;
      boundaries.push(
        Object.freeze({
          tableId: table.id,
          role: before.rowId === after.rowId ? 'cell' : 'row',
          beforeCellId: before.id,
          afterCellId: after.id,
          extent:
            start !== undefined && end !== undefined && start <= end
              ? Object.freeze({ start, end })
              : null,
        })
      );
    }
  }
  return {
    success: true,
    projection: Object.freeze({
      ...projection,
      tables: Object.freeze(
        [...tables.values()].map((table) =>
          Object.freeze({ ...table, rowIds: Object.freeze([...table.rowIds]) })
        )
      ),
      cells: Object.freeze(
        [...cells.values()].map((cell) =>
          Object.freeze({
            ...cell,
            paragraphIds: Object.freeze([...cell.paragraphIds]),
            sources: Object.freeze(cell.sources),
          })
        )
      ),
      boundaries: Object.freeze(boundaries),
    }),
  };
}

/** Generated table separators and escape interiors never select a cell's native text. */
export function tableCellAt(
  projection: TableProjection,
  offset: number
): ProjectedTableCell | null {
  const point = markdownPointAt(projection, offset);
  if (!point || point.sourceScope !== 'body') return null;
  return projection.cells.find((cell) => cell.paragraphIds.includes(point.paragraphId)) ?? null;
}
