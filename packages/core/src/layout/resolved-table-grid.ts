import type { OoxmlNode } from '@docx-editor.dev/core/store';

interface RowGrid {
  readonly cells: readonly OoxmlNode[];
  readonly starts: readonly number[];
  readonly spans: readonly number[];
  readonly gridColumns: number;
}
interface ResolvedCellGrid {
  readonly resolvedGridColumn: number;
  readonly resolvedGridSpan: number;
  readonly resolvedGridColumnCount: number;
}
/** Semantic grid after omitted revision cells/rows, without changing physical layout coordinates. */
export function resolvedTableGrid(
  rows: readonly RowGrid[],
  hasOmissions: boolean
): ReadonlyMap<string, ResolvedCellGrid> {
  const result = new Map<string, ResolvedCellGrid>();
  if (!hasOmissions) return result;
  const boundaries = new Set<number>([0]);
  for (const row of rows) {
    boundaries.add(row.gridColumns);
    for (let index = 0; index < row.starts.length; index++) {
      boundaries.add(row.starts[index]!);
      boundaries.add(row.starts[index]! + row.spans[index]!);
    }
  }
  const sorted = [...boundaries].sort((a, b) => a - b),
    positions = new Map(sorted.map((column, index) => [column, index]));
  for (const row of rows) {
    let index = 0;
    for (const cell of row.cells) {
      if (cell.kind !== 'tableCell') continue;
      const start = row.starts[index]!,
        span = row.spans[index++]!;
      const column = positions.get(start)!,
        width = positions.get(start + span)! - column;
      result.set(cell.id, {
        resolvedGridColumn: column,
        resolvedGridSpan: width,
        resolvedGridColumnCount: sorted.length - 1,
      });
    }
  }
  return result;
}
