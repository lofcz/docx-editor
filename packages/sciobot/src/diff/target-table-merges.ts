import { planInsertedCellAtoms } from './inserted-cell-atoms';
import { targetNewCellProperties, type NewCellPropertyBinding } from './new-cell-properties';
import { applyTreeOp, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { TargetAstNode } from '../projection/parse-target';
import type { LogicalBlock, LogicalGroup } from './logical-block-types';
import type { CanonicalEdit } from './plan';
import type { TargetTableContext } from './target-table';
import type { TargetTablePropertyCodec } from './target-table-property-codec';
import { planTableCellEdits } from './table-cells';
import { planTableMerge } from './table-merges';
import { blankCellTarget, sameTargetCellContent } from './target-table-content';
import { insertedTargetCellFragments } from './target-table-fragments';
import {
  targetSlotKey as slot,
  type SourceSlotGrid,
  type TargetGeometryRequest,
  type TargetSlotRegion,
} from './target-table-spans';

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

/**
 * Lower target merge/split requests after content planning. Native merges move every filled
 * cell's blocks into the head in reading order, so covered cells whose content the target keeps
 * as an exact suffix stay intact; earlier covered cells are emptied and the head receives the rest.
 * Splits keep the head cell and its content; new cells carry only their target fragments.
 */
export function planTargetGeometry(
  context: TargetTableContext,
  codec: TargetTablePropertyCodec,
  requests: readonly TargetGeometryRequest[],
  grid: SourceSlotGrid | null,
  initial: OoxmlPart
): { readonly success: true; readonly operations: readonly CanonicalEdit[] } | NativeEditFailure {
  if (!requests.length) return { success: true, operations: [] };
  if (!grid) return fail('Table merges need native merge geometry.');
  const { input, source, before, after, beforePath, afterPath } = context,
    revision = source.index.revision,
    tableId = context.binding.table!.tableId,
    oldRows = before.children as readonly LogicalGroup[],
    newRows = after.children as readonly LogicalGroup[],
    operations: CanonicalEdit[] = [];
  const cellAt = (rows: readonly LogicalGroup[], region: { row: number; column: number }) =>
    rows[region.row]!.children[region.column] as LogicalGroup;
  const visit = (
    head: TargetSlotRegion,
    next: readonly LogicalBlock[],
    target: TargetSlotRegion
  ): NativeEditFailure | null => {
    const old = cellAt(oldRows, head),
      oldPath = [...beforePath, head.row, head.column],
      newPath = [...afterPath, target.row, target.column],
      cellId = grid.cellIds.get(slot(head.row, head.column));
    if (!cellId) return fail('A merged region has no exact native head cell.');
    if (sameTargetCellContent(context, old.children, next, oldPath)) return null;
    const owner = context.bindings.get(oldPath.join('/'))?.table;
    if (!owner?.rowId || owner.cellId !== cellId)
      return fail('A merged region head has no exact canonical table/row binding.');
    const checked = planTableCellEdits(
      {
        revision,
        part: input.pkg.parts.get(source.index.partName)!,
        projection: input.projection,
        displayMode: input.projection.displayMode!,
        tableId,
        rowId: owner.rowId,
        cellId,
      },
      []
    );
    if (!checked.success) return checked;
    return context.visitCell(
      old.children,
      blankCellTarget(old.children, next, newPath),
      oldPath,
      newPath,
      cellId
    );
  };
  let part = initial;
  const push = (edit: CanonicalEdit, operation: Parameters<typeof applyTreeOp>[1]) => {
    const applied = applyTreeOp(part, operation);
    if (!applied.ok) return fail(`Combined merge/split refused: ${applied.reason}.`);
    part = applied.part;
    operations.push(edit);
    return null;
  };
  for (const request of requests) {
    if (request.kind === 'merge') {
      const anchor = cellAt(newRows, request.region);
      if (codec.key(anchor) !== codec.key(cellAt(oldRows, request.heads[0]!)))
        return fail('Changed merged cell metadata needs explicit native property intentions.');
      const target = anchor.children;
      // Keep a trailing run of covered cells whose complete content survives as the target suffix.
      let end = target.length,
        retaining = true;
      const emptied: TargetSlotRegion[] = [];
      for (let index = request.heads.length - 1; index >= 1; index--) {
        const head = request.heads[index]!,
          blocks = cellAt(oldRows, head).children;
        if (!blocks.length) continue;
        if (
          retaining &&
          blocks.length <= end &&
          sameTargetCellContent(context, blocks, target.slice(end - blocks.length, end), [
            ...beforePath,
            head.row,
            head.column,
          ])
        ) {
          end -= blocks.length;
          continue;
        }
        retaining = false;
        emptied.push(head);
      }
      for (const head of emptied) {
        const error = visit(head, [], request.region);
        if (error) return error;
      }
      const error = visit(request.heads[0]!, target.slice(0, end), request.region);
      if (error) return error;
      const cellIds = request.heads.map((head) => grid.cellIds.get(slot(head.row, head.column))!);
      const planned = planTableMerge({ revision, part, tableId }, { kind: 'merge', cellIds });
      if (!planned.success) return planned;
      for (const edit of planned.plan.operations) {
        if (edit.kind !== 'table') return fail('Unexpected native merge intention.');
        const refused = push(edit, edit.operation);
        if (refused) return refused;
      }
      continue;
    }
    const [first, ...rest] = request.parts,
      cellId = grid.cellIds.get(slot(request.region.row, request.region.column))!;
    const error = visit(request.region, cellAt(newRows, first!).children, first!);
    if (error) return error;
    const cells: (readonly TargetAstNode[])[] = [];
    const cellProperties: NewCellPropertyBinding[] = [];
    for (const part of rest) {
      const cell = cellAt(newRows, part);
      if (codec.key(cell) !== codec.key(cellAt(oldRows, part)) && cell.key !== 'covered')
        return fail('New split cell properties need deferred property destinations.');
      const fragments = insertedTargetCellFragments(
        cell,
        context.targetAst,
        context.astOf,
        undefined,
        codec
      );
      const properties = targetNewCellProperties(
        codec,
        after,
        newRows[part.row]!,
        cell,
        part.column
      );
      if (Object.keys(properties).length)
        cellProperties.push({ cellIndex: cells.length, properties });
      if (!fragments.success) return fragments;
      cells.push(fragments.fragments);
    }
    const atomPlan = planInsertedCellAtoms(input.pkg, cells, [], context.allocateBlockId);
    if (!atomPlan.success) return atomPlan;
    const split = {
      kind: 'split' as const,
      cellId,
      rows: request.rows,
      columns: request.columns,
      columnDistribution: 'existing-grid' as const,
    };
    const planned = planTableMerge({ revision, part, tableId }, split);
    if (!planned.success) return planned;
    const operation = planned.plan.operations[0];
    if (operation?.kind !== 'table') return fail('Unexpected native split intention.');
    const refused =
      cellProperties.length || cells.some((blocks) => blocks.length)
        ? push(
            {
              kind: 'split-cell',
              partName: operation.partName,
              intent: {
                tableId,
                cellId,
                rows: request.rows,
                columns: request.columns,
                cells,
                payloads: [],
                ...(atomPlan.atoms.length ? { atoms: atomPlan.atoms } : {}),
                ...(cellProperties.length ? { cellProperties } : {}),
              },
            },
            operation.operation
          )
        : push(operation, operation.operation);
    if (refused) return refused;
  }
  return { success: true, operations };
}
