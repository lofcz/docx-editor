import { isInsertedCellAtoms } from './inserted-cell-atoms';
import type { InsertedInlineAtom } from './inserted-atoms';
import { isNewCellPropertyBindings, type NewCellPropertyBinding } from './new-cell-properties';
import { readTableStructure } from '@docx-editor.dev/core/layout';
import { findNode, type OoxmlPart, type TreeDocOp } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { TargetAstNode } from '../projection/parse-target';
import { isInsertedCellFragments } from './inserted-cell-fragments';
import { type SourceSlotGrid, type TargetSlotRegion } from './target-table-spans';

/** Resolve coordinates after earlier native topology allocation, never from preview node IDs. */
export type TableGeometryIntent = {
  readonly tableId: string;
  readonly region: TargetSlotRegion;
} & (
  | { readonly kind: 'merge' }
  | {
      readonly kind: 'split';
      readonly rows: number;
      readonly columns: number;
      readonly cells: readonly (readonly TargetAstNode[])[];
      readonly atoms?: readonly InsertedInlineAtom[];
      readonly cellProperties?: readonly NewCellPropertyBinding[];
    }
);
const slot = (row: number, column: number) => `${row}/${column}`;
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const integer = (value: unknown, minimum: number) =>
  Number.isSafeInteger(value) && Number(value) >= minimum && Number(value) <= 63;
export function isTableGeometryIntent(value: unknown): value is TableGeometryIntent {
  if (
    !record(value) ||
    typeof value.tableId !== 'string' ||
    !value.tableId ||
    !record(value.region)
  )
    return false;
  const r = value.region;
  if (
    !Object.keys(r).every((key) => ['row', 'column', 'rows', 'columns'].includes(key)) ||
    !(
      Number.isSafeInteger(r.row) &&
      Number(r.row) >= 0 &&
      Number.isSafeInteger(Number(r.row) + Number(r.rows))
    ) ||
    !integer(r.column, 0) ||
    !integer(r.rows, 1) ||
    !integer(r.columns, 1) ||
    Number(r.column) + Number(r.columns) > 63
  )
    return false;
  if (value.kind === 'merge')
    return (
      Object.keys(value).every((key) => ['tableId', 'region', 'kind'].includes(key)) &&
      Number(r.rows) * Number(r.columns) > 1
    );
  return (
    value.kind === 'split' &&
    Object.keys(value).every((key) =>
      ['tableId', 'region', 'kind', 'rows', 'columns', 'cells', 'atoms', 'cellProperties'].includes(
        key
      )
    ) &&
    integer(value.rows, 1) &&
    integer(value.columns, 1) &&
    Number(value.rows) * Number(value.columns) > 1 &&
    Array.isArray(value.cells) &&
    value.cells.length === Number(value.rows) * Number(value.columns) - 1 &&
    isInsertedCellFragments(value.cells, [], true) &&
    isInsertedCellAtoms(value.cells as TargetAstNode[][], [], value.atoms) &&
    (!Object.hasOwn(value, 'cellProperties') ||
      isNewCellPropertyBindings(value.cells as TargetAstNode[][], value.cellProperties))
  );
}

/** Pure native Proposed geometry, including omitted rows/cells and virtual merges. */
export function nativeTableSlotGrid(part: OoxmlPart, tableId: string): SourceSlotGrid | null {
  const node = findNode(part, tableId);
  if (node?.kind !== 'table') return null;
  const table = readTableStructure(node, 500, 0, undefined, 'proposed');
  if (!table) return null;
  const heads: { row: number; column: number }[][] = [],
    cellIds = new Map<string, string>(),
    regions = new Map<string, TargetSlotRegion>();
  for (const [row, entry] of table.rows.entries()) {
    heads.push([]);
    for (const cell of entry.cells) {
      const column = cell.resolvedGridColumn ?? cell.logicalGridColumn ?? cell.gridColumn,
        span = cell.resolvedGridSpan ?? cell.gridSpan,
        head = cell.vMergeContinue ? heads[row - 1]?.[column] : { row, column };
      if (!head) return null;
      const key = slot(head.row, head.column);
      if (!cell.vMergeContinue) cellIds.set(key, cell.id);
      for (let c = column; c < column + span; c++) heads[row]![c] = head;
      const previous = regions.get(key);
      regions.set(key, {
        row: head.row,
        column: head.column,
        rows: row - head.row + 1,
        columns: previous?.columns ?? span,
      });
    }
  }
  const width = heads[0]?.length ?? 0;
  if (
    !width ||
    heads.some(
      (row) =>
        row.length !== width ||
        Array.from({ length: width }, (_, c) => row[c]).some((cell) => !cell)
    )
  )
    return null;
  for (const [key, region] of regions)
    for (let r = region.row; r < region.row + region.rows; r++)
      for (let c = region.column; c < region.column + region.columns; c++) {
        const head = heads[r]?.[c];
        if (!head || slot(head.row, head.column) !== key) return null;
      }
  return { heads, cellIds, regions };
}
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
export function resolveTableGeometry(
  part: OoxmlPart,
  intent: TableGeometryIntent
):
  | { readonly success: true; readonly operation: TreeDocOp; readonly cellId: string }
  | NativeEditFailure {
  const grid = nativeTableSlotGrid(part, intent.tableId),
    r = intent.region;
  if (!grid) return fail('The staged table has no complete native slot grid.');
  if (r.row + r.rows > grid.heads.length || r.column + r.columns > (grid.heads[0]?.length ?? 0))
    return fail('The staged table region exceeds its native grid.');
  const regions = new Map<string, TargetSlotRegion>();
  for (let row = r.row; row < r.row + r.rows; row++)
    for (let c = r.column; c < r.column + r.columns; c++) {
      const head = grid.heads[row]?.[c];
      if (!head) return fail('The staged table region no longer exists.');
      const key = slot(head.row, head.column),
        region = grid.regions.get(key)!;
      if (
        region.row < r.row ||
        region.column < r.column ||
        region.row + region.rows > r.row + r.rows ||
        region.column + region.columns > r.column + r.columns
      )
        return fail('The staged region crosses a native merged cell.');
      regions.set(key, region);
    }
  const ids = [...regions.keys()].map((key) => grid.cellIds.get(key)!);
  if (intent.kind === 'merge')
    return ids.length < 2
      ? fail('The staged merge needs at least two native cells.')
      : {
          success: true,
          cellId: ids[0]!,
          operation: { op: 'mergeTableCells', tableId: intent.tableId, cellIds: ids },
        };
  if (ids.length !== 1 || r.rows % intent.rows || r.columns % intent.columns)
    return fail('The staged split needs one complete native cell and exact grid cuts.');
  return {
    success: true,
    cellId: ids[0]!,
    operation: {
      op: 'splitTableCell',
      tableId: intent.tableId,
      cellId: ids[0]!,
      rows: intent.rows,
      columns: intent.columns,
      columnDistribution: 'existing-grid',
    },
  };
}
/** Actual allocated split heads, in target reading order, excluding the retained head. */
export function geometrySplitCells(
  part: OoxmlPart,
  intent: TableGeometryIntent,
  headId: string
): readonly string[] | null {
  const grid = nativeTableSlotGrid(part, intent.tableId);
  if (!grid) return null;
  const r = intent.region;
  return [...grid.regions]
    .filter(
      ([, own]) =>
        own.row >= r.row &&
        own.column >= r.column &&
        own.row + own.rows <= r.row + r.rows &&
        own.column + own.columns <= r.column + r.columns
    )
    .sort(([, a], [, b]) => a.row - b.row || a.column - b.column)
    .map(([key]) => grid.cellIds.get(key)!)
    .filter((id) => id !== headId);
}
