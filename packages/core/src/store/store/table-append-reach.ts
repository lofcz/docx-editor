import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { readEditableTableTopology } from './tree-op-table-topology.ts';
import { isWmlElement } from './tree-op-table-shared.ts';
import type { TreeDocOp } from './tree-op-types.ts';

/** A fresh row or cell beyond the last edge leaves the existing cells and their controls in place. */
export function appendsOutsideExistingCells(part: OoxmlPart, op: TreeDocOp): boolean {
  if (op.op !== 'insertTableRow' && op.op !== 'insertTableColumn') return false;
  const result = readEditableTableTopology(part.root, op.tableId);
  if (!result.ok) return false;
  const { table, rows, gridColumns, hasMerge } = result.topology;
  // Wrapped rows/cells need structural protection. Admit the ordinary, directly authored grid.
  const space = (node: OoxmlNode) => node.kind === 'textValue' && !node.value.trim();
  if (
    hasMerge ||
    table.children.some(
      (node) =>
        !space(node) &&
        node.kind !== 'tableRow' &&
        !isWmlElement(node, 'tblPr') &&
        !isWmlElement(node, 'tblGrid')
    )
  )
    return false;
  if (
    rows.some(({ row }) =>
      row.children.some(
        (node) => !space(node) && node.kind !== 'tableCell' && !isWmlElement(node, 'trPr')
      )
    )
  )
    return false;
  if (op.op === 'insertTableRow') return op.where === 'below' && rows.at(-1)?.row.id === op.rowId;
  if (op.where !== 'right') return false;
  return 'gridColumnId' in op
    ? gridColumns.at(-1)?.id === op.gridColumnId
    : rows[0]?.cells.at(-1)?.id === op.referenceCellId;
}
