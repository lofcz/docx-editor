import { applyTreeOp, findNode, withPart, type OoxmlPackage } from '@docx-editor.dev/core/store';
import type { InsertedTableRowIntent } from '../../src/diff/inserted-row-types';
import type { CanonicalEditPlan } from '../../src/diff/plan';
import { materializeNewCellFragments } from './inserted-fragment-oracle';

import { rowCells, tableRows } from './native-table-oracle';
export { nodeChildren, namedChild, tableRows, rowCells } from './native-table-oracle';

/**
 * Test-only fragment codec for the acceptance fixture. This is not the application executor.
 * Native row/fragment operations allocate all document IDs after planning.
 */
export function materializeInsertedRow(
  input: OoxmlPackage,
  plan: CanonicalEditPlan
): {
  pkg: OoxmlPackage;
  rowId: string;
} {
  const edit = plan.operations[0];
  if (plan.operations.length !== 1 || edit?.kind !== 'insert-row')
    throw new Error('Expected one rich row');
  const intent: InsertedTableRowIntent = edit.intent;
  let pkg = input,
    part = pkg.parts.get(edit.partName)!;
  const old = new Set(tableRows(part, intent.tableId).map((node) => node.id));
  const inserted = applyTreeOp(part, {
    op: 'insertTableRow',
    tableId: intent.tableId,
    rowId: intent.rowId,
    where: intent.where,
  });
  if (!inserted.ok) throw new Error(inserted.reason);
  part = inserted.part;
  const row = tableRows(part, intent.tableId).find((node) => !old.has(node.id))!;
  if (intent.properties.heightTwips !== undefined) {
    const sized = applyTreeOp(part, {
      op: 'setTableRowHeight',
      tableId: intent.tableId,
      rowId: row.id,
      heightTwips: intent.properties.heightTwips,
    });
    if (!sized.ok) throw new Error(sized.reason);
    part = sized.part;
  }
  pkg = withPart(pkg, part);
  const cells = rowCells(findNode(part, row.id)!);
  const result = materializeNewCellFragments(
    pkg,
    edit.partName,
    cells.map((cell, index) => ({ cellId: cell.id, blocks: intent.cells[index]! })),
    intent.payloads
  );
  return { pkg: result, rowId: row.id };
}
