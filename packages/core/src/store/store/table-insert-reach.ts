import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { readEditableTableTopology } from './tree-op-table-topology.ts';
import { isWmlElement, wmlChildNamed } from './tree-op-table-shared.ts';
import type { TreeDocOp } from './tree-op-types.ts';

/** A blank sibling does not edit a control inside an existing cell paragraph. */
export function insertsOutsideCellContents(part: OoxmlPart, op: TreeDocOp): boolean {
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
  const appends =
    op.op === 'insertTableRow'
      ? op.where === 'below' && rows.at(-1)?.row.id === op.rowId
      : op.where === 'right' &&
        ('gridColumnId' in op
          ? gridColumns.at(-1)?.id === op.gridColumnId
          : rows[0]?.cells.at(-1)?.id === op.referenceCellId);
  if (appends) return true;
  // The inserter copies cell and paragraph formatting into a new, empty sibling.
  // Interior edits of wrapped blocks, skipped cells, and tracked structure keep their
  // structural check. A lock around the table is still checked at the table itself.
  if (op.revision) return false;
  return rows.every(({ row, cells }) => {
    const props = wmlChildNamed(row, 'trPr');
    return (
      cells.length === gridColumns.length &&
      !(props && (wmlChildNamed(props, 'gridBefore') || wmlChildNamed(props, 'gridAfter'))) &&
      cells.every((cell) =>
        cell.children.every(
          (node) => space(node) || node.kind === 'paragraph' || isWmlElement(node, 'tcPr')
        )
      )
    );
  });
}
