import { applyTreeOp, findNode, withPart, type OoxmlPackage } from '@docx-editor.dev/core/store';
import {
  insertedColumnIndex,
  insertedColumnOperation,
  insertedColumnWidthOperation,
  nativeGridColumns,
} from '../../src/diff/inserted-column-native';
import type { CanonicalEditPlan } from '../../src/diff/plan';
import { materializeNewCellFragments } from './inserted-fragment-oracle';
import { rowCells } from './native-table-oracle';

/** Test-only immutable materialization; A01/R09 must supply guards, transactions, and review. */
export function materializeInsertedColumn(
  input: OoxmlPackage,
  plan: CanonicalEditPlan
): {
  pkg: OoxmlPackage;
  gridColumnId: string;
  cellIds: ReadonlyMap<string, string>;
} {
  const edit = plan.operations[0];
  if (plan.operations.length !== 1 || edit?.kind !== 'insert-column')
    throw new Error('Expected one rich column');
  const intent = edit.intent,
    before = input.parts.get(edit.partName)!;
  const inserted = applyTreeOp(before, insertedColumnOperation(intent));
  if (!inserted.ok) throw new Error(inserted.reason);
  let part = inserted.part;
  const index = insertedColumnIndex(before, part, intent);
  if (index === null) throw new Error('Missing inserted native grid column');
  const gridColumnId = nativeGridColumns(part, intent.tableId)[index]!.id;
  const sizing = insertedColumnWidthOperation(before, part, intent);
  if (!sizing.ok) throw new Error(sizing.reason);
  if (sizing.operation) {
    const sized = applyTreeOp(part, sizing.operation);
    if (!sized.ok) throw new Error(sized.reason);
    part = sized.part;
  }
  const cellIds = new Map<string, string>();
  const cells = intent.cells.map((entry) => {
    const old = new Set(rowCells(findNode(before, entry.rowId)!).map((cell) => cell.id));
    const created = rowCells(findNode(part, entry.rowId)!).filter((cell) => !old.has(cell.id));
    if (created.length > 1 || (!created.length && entry.blocks.length))
      throw new Error('Ambiguous fixture destination');
    const cellId = created[0]?.id;
    if (cellId) cellIds.set(entry.rowId, cellId);
    return { ...(cellId ? { cellId } : {}), blocks: entry.blocks };
  });
  return {
    pkg: materializeNewCellFragments(withPart(input, part), edit.partName, cells, intent.payloads),
    gridColumnId,
    cellIds,
  };
}
