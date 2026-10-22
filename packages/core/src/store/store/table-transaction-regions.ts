import { WML_NAMESPACE_URI, type OoxmlElement, type OoxmlNode } from '../package/ooxml-tree.ts';
import type { TreeDocOp } from './tree-op-types.ts';
import { child, kids } from './table-transaction-state.ts';

interface Region {
  top: number;
  bottom: number;
  left: number;
  right: number;
}
const value = (node: OoxmlNode | undefined, name: string) =>
  node && node.kind !== 'textValue'
    ? node.attributes.find(
        (attr) => attr.namespaceUri === WML_NAMESPACE_URI && attr.localName === name
      )?.value
    : undefined;
/** Bounds of virtual merges this frame already proposed, in the still-physical pending grid. */
export function pendingTableMergeRegions(
  table: OoxmlElement,
  owned: ReadonlyMap<string, string>
): {
  regions: Region[];
  cells: ReadonlyMap<string, { row: number; left: number; right: number }>;
  rows: readonly string[];
} {
  const regions: Region[] = [],
    cells = new Map<string, { row: number; left: number; right: number }>(),
    rows = kids(table).filter((row) => row.kind === 'tableRow'),
    vertical = new Map<number, Region>();
  rows.forEach((row, index) => {
    let column = Number(value(child(child(row, 'trPr'), 'gridBefore'), 'val') ?? 0),
      previous: { left: number; right: number } | undefined;
    for (const cell of kids(row))
      if (cell.kind === 'tableCell') {
        const pr = child(cell, 'tcPr'),
          span = Math.max(1, Number(value(child(pr, 'gridSpan'), 'val') ?? 1)),
          left = column,
          right = left + span - 1;
        column += span;
        cells.set(cell.id, { row: index, left, right });
        const removed = child(pr, 'cellDel');
        if (removed) {
          if (owned.get(removed.id) === 'mergeTableCells' && previous) {
            regions.push({ top: index, bottom: index, left: previous.left, right });
            previous.right = right;
          }
          continue;
        }
        previous = { left, right };
        const merge = child(pr, 'cellMerge');
        if (!merge || owned.get(merge.id) !== 'mergeTableCells') continue;
        if (value(merge, 'vMerge') === 'rest') {
          const region = { top: index, bottom: index, left, right };
          vertical.set(left, region);
          regions.push(region);
        } else {
          const head = vertical.get(left);
          if (head) head.bottom = index;
        }
      }
  });
  return { regions, cells, rows: rows.map((row) => row.id) };
}
/**
 * Current target programs create topology before their disjoint geometry regions. A later edit
 * inside a virtual merge needs a new region plan; do not reinterpret its old physical slots.
 */
export function overlapsPendingMerge(
  table: OoxmlElement,
  op: TreeDocOp,
  owned: ReadonlyMap<string, string>
): boolean {
  const { regions, cells, rows } = pendingTableMergeRegions(table, owned);
  if (!regions.length) return false;
  if (op.op === 'splitTableCell') {
    const cell = cells.get(op.cellId);
    return (
      !!cell &&
      regions.some(
        (region) =>
          cell.row >= region.top &&
          cell.row <= region.bottom &&
          cell.left <= region.right &&
          cell.right >= region.left
      )
    );
  }
  if (op.op === 'deleteTableRow') {
    const row = rows.indexOf(op.rowId);
    return regions.some((region) => row >= region.top && row <= region.bottom);
  }
  if (op.op === 'insertTableRow') {
    const index = rows.indexOf(op.rowId) + (op.where === 'above' ? 0 : 1);
    return regions.some((region) => index > region.top && index <= region.bottom);
  }
  if (op.op === 'insertTableColumn' || op.op === 'deleteTableColumn') {
    const grid = kids(child(table, 'tblGrid')).filter(
      (column) => column.kind !== 'textValue' && column.localName === 'gridCol'
    );
    const anchor =
      'gridColumnId' in op && op.gridColumnId
        ? grid.findIndex((column) => column.id === op.gridColumnId)
        : 'referenceCellId' in op && op.referenceCellId
          ? (cells.get(op.referenceCellId)?.left ?? -1)
          : -1;
    if (anchor < 0) return true;
    if (op.op === 'deleteTableColumn')
      return regions.some((region) => anchor >= region.left && anchor <= region.right);
    const insertion = anchor + (op.where === 'left' ? 0 : 1);
    return regions.some((region) => insertion > region.left && insertion <= region.right);
  }
  return false;
}
