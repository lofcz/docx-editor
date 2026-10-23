import { targetNewCellProperties, type NewCellPropertyBinding } from './new-cell-properties';
import type { InsertedCellListBinding } from './inserted-cell-lists';
import { readTableStructure } from '@docx-editor.dev/core/layout';
import { planTargetShapeGeometry, type AllocatedTargetCell } from './target-table-shape';
import { rowCellIds, tableRowIds } from './fragment-seeds';
import { applyTreeOp, findNode, withPart } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { TargetAstNode } from '../projection/parse-target';
import type { BoundLogicalSource, BoundSourceBlock } from './bound-source-types';
import type { LogicalBlock, LogicalGroup } from './logical-block-types';
import type { TargetSourceInput } from './target-source-types';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';
import { planTableCellEdits } from './table-cells';
import { planTableRow } from './table-rows';
import { planInsertedTableRow } from './inserted-table-rows';
import { targetAstKey } from './target-key';
import { SIBLING_BLOCK_LIMITS } from './sibling-blocks';
import { alignTargetSequence } from './target-sequence';
import { alignTargetColumns, planTargetColumns } from './target-table-columns';
import { prepareTargetTableProperties } from './target-table-property-codec';
import { planTargetTableProperties } from './target-table-properties';
import { blankCellTarget, sameTargetCellContent } from './target-table-content';
import { insertedTargetCellFragments } from './target-table-fragments';
import {
  expandTargetSpans,
  sourceSlotGrid,
  targetGeometryRequests,
  type TargetGeometryRequest,
} from './target-table-spans';
import { planTargetGeometry } from './target-table-merges';

export interface TargetTableContext {
  readonly allocateBlockId?: () => string;
  readonly input: TargetSourceInput;
  readonly source: BoundLogicalSource;
  readonly before: LogicalGroup;
  readonly after: LogicalGroup;
  readonly beforePath: readonly number[];
  readonly afterPath: readonly number[];
  readonly binding: BoundSourceBlock;
  readonly bindings: ReadonlyMap<string, BoundSourceBlock>;
  readonly astOf: (block: LogicalBlock) => TargetAstNode;
  readonly targetAst: TargetAstNode;
  readonly same: (before: LogicalBlock, after: LogicalBlock) => boolean;
  readonly visitCell: (
    before: readonly LogicalBlock[],
    after: readonly LogicalBlock[],
    beforePath: readonly number[],
    afterPath: readonly number[],
    cellId: string
  ) => NativeEditFailure | null;
}
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

function groups(table: LogicalGroup): string {
  return JSON.stringify(
    table.rowGroups.map((group) => ({
      header: targetAstKey({ ...group.header, type: 'row-group' }),
      start: group.start,
      end: group.end,
    }))
  );
}

/** Retain native table/row/cell owners. Only new rows carry target fragments. */
export function planTargetTable(original: TargetTableContext): EditPlanResult {
  if (original.after.kind !== 'table') return fail('The retained table needs a target table.');
  const expansion = expandTargetSpans(original.after);
  if (!expansion.success) return expansion;
  const context: TargetTableContext = { ...original, after: expansion.expanded.table };
  const { input, source, before, after, beforePath, afterPath, binding, bindings } = context,
    part = input.pkg.parts.get(source.index.partName)!,
    revision = source.index.revision,
    tableId = binding.table?.tableId,
    native = tableId ? findNode(part, tableId) : null;
  if (
    before.kind !== 'table' ||
    after.kind !== 'table' ||
    !tableId ||
    native?.kind !== 'table' ||
    binding.table?.role !== 'table'
  )
    return fail('The retained table needs one exact canonical table owner.');
  const properties = prepareTargetTableProperties(before, after);
  if (!properties.success) return properties;
  const codec = properties.codec;
  if (codec.key(before) !== codec.key(after) || groups(before) !== groups(after))
    return fail('Changed table metadata or row groups need explicit native property intentions.');
  if (!before.children.length || !after.children.length)
    return fail('Empty table topology needs explicit native seed or removal intentions.');
  const oldRows = before.children as readonly LogicalGroup[],
    newRows = after.children as readonly LogicalGroup[],
    width = oldRows[0]!.children.length,
    newWidth = newRows[0]!.children.length;
  if (
    oldRows.length > SIBLING_BLOCK_LIMITS.maxSiblings ||
    newRows.length > SIBLING_BLOCK_LIMITS.maxSiblings
  )
    return fail('Table row matching exceeds its sibling limit.');
  if (
    !width ||
    !newWidth ||
    oldRows.some(
      (row) =>
        row.kind !== 'row' ||
        row.children.length !== width ||
        row.children.some((cell) => cell.kind !== 'cell')
    ) ||
    newRows.some(
      (row) =>
        row.kind !== 'row' ||
        row.children.length !== newWidth ||
        row.children.some((cell) => cell.kind !== 'cell')
    )
  )
    return fail('Irregular table slot counts need explicit native grid intentions.');
  const rowIds = oldRows.map(
    (_, index) => bindings.get([...beforePath, index].join('/'))?.table?.rowId
  );
  if (rowIds.some((id) => !id) || new Set(rowIds).size !== rowIds.length)
    return fail('The source table rows need distinct canonical bindings.');
  // Native merges and explicit target spans: compare partitions slot by slot.
  const grid = sourceSlotGrid(bindings, beforePath, oldRows.length, width),
    changedShape = oldRows.length !== newRows.length || width !== newWidth;
  let requests: readonly TargetGeometryRequest[] = [];
  if (grid && oldRows.length === newRows.length && width === newWidth) {
    const compared = targetGeometryRequests(grid, expansion.expanded);
    if (!compared.success) return compared;
    requests = compared.requests;
  }
  const positional = (count: number) => ({
    success: true as const,
    alignment: {
      pairs: Array.from({ length: count }, (_, index) => ({ before: index, after: index })),
      gaps: [],
    },
  });
  const handled = new Set<string>();
  for (const request of requests)
    for (let r = request.region.row; r < request.region.row + request.region.rows; r++)
      for (let c = request.region.column; c < request.region.column + request.region.columns; c++)
        handled.add(`${r}/${c}`);
  const alignedColumns = requests.length
    ? positional(width)
    : alignTargetColumns(oldRows, newRows, codec.contentKey);
  if (!alignedColumns.success) return alignedColumns;
  const projectedRows = (rows: readonly LogicalGroup[], side: 'before' | 'after') =>
    rows.map((row) =>
      JSON.stringify([
        codec.key(row),
        alignedColumns.alignment.pairs.map((pair) => codec.contentKey(row.children[pair[side]]!)),
      ])
    );
  const alignedRows = requests.length
    ? positional(oldRows.length)
    : alignTargetSequence(projectedRows(oldRows, 'before'), projectedRows(newRows, 'after'));
  if (!alignedRows.success) return alignedRows;
  const { pairs, gaps } = alignedRows.alignment;
  const columns = planTargetColumns(
    context,
    alignedColumns.alignment,
    pairs,
    rowIds as readonly string[],
    codec
  );
  if (!columns.success) return columns;
  let operations: CanonicalEdit[] = [...columns.operations];
  const allocated: AllocatedTargetCell[] = [...columns.allocated];
  let preview = columns.part;
  for (const gap of [...gaps].reverse()) {
    for (let index = gap.beforeEnd - 1; index >= gap.beforeStart; index--) {
      const result = planTableRow(
        { revision, part: preview, tableId },
        { kind: 'delete', rowId: rowIds[index]! }
      );
      if (!result.success) return result;
      for (const edit of result.plan.operations) {
        if (edit.kind !== 'table') return fail('Unexpected native row deletion intent.');
        const removed = applyTreeOp(preview, edit.operation);
        if (!removed.ok) return fail(`Combined row deletion refused: ${removed.reason}.`);
        preview = removed.part;
        operations.push(edit);
      }
    }
    const anchor = rowIds[gap.beforeStart] ?? rowIds[gap.beforeStart - 1];
    const where = rowIds[gap.beforeStart] ? 'above' : 'below',
      insertedIndices = Array.from(
        { length: gap.afterEnd - gap.afterStart },
        (_, index) => gap.afterStart + index
      );
    if (where === 'below') insertedIndices.reverse();
    for (const index of insertedIndices) {
      const row = newRows[index]!;
      if (codec.key(row) !== targetAstKey({ type: 'row' }))
        return fail('Inserted row metadata needs explicit native properties.');
      const inserted = applyTreeOp(preview, {
        op: 'insertTableRow',
        tableId,
        rowId: anchor!,
        where,
      });
      if (!inserted.ok) return fail(`Combined row insertion refused: ${inserted.reason}.`);
      const old = new Set(tableRowIds(preview, tableId));
      const fresh = tableRowIds(inserted.part, tableId).find((id) => !old.has(id));
      if (!fresh) return fail('The topology preview did not allocate one native row.');
      const structure = readTableStructure(
        findNode(inserted.part, tableId)!,
        500,
        0,
        undefined,
        'proposed'
      );
      const nativeRow = structure?.rows.find((row) => row.id === fresh);
      if (!nativeRow) return fail('The inserted row has no native slot geometry.');
      const cells: (readonly TargetAstNode[])[] = [];
      const lists: InsertedCellListBinding[] = [];
      const cellProperties: NewCellPropertyBinding[] = [];
      for (const nativeCell of nativeRow.cells) {
        const cell = row.children[
          nativeCell.logicalGridColumn ?? nativeCell.gridColumn
        ] as LogicalGroup;
        if (!cell) return fail('The inserted row extends beyond its target slots.');
        const fragments = insertedTargetCellFragments(
          cell,
          context.targetAst,
          context.astOf,
          {
            revision,
            pkg: input.pkg,
            partName: part.name,
          },
          codec
        );
        if (!fragments.success) return fragments;
        if (!nativeCell.vMergeContinue)
          lists.push(
            ...fragments.lists.map((binding) => ({
              ...binding,
              path: [cells.length, ...binding.path],
            }))
          );
        const properties = targetNewCellProperties(
          codec,
          after,
          row,
          cell,
          nativeCell.logicalGridColumn ?? nativeCell.gridColumn
        );
        if (Object.keys(properties).length)
          cellProperties.push({ cellIndex: cells.length, properties });
        cells.push(nativeCell.vMergeContinue ? [] : fragments.fragments);
      }
      const result = planInsertedTableRow(
        {
          revision,
          pkg: withPart(input.pkg, preview),
          partName: part.name,
          tableId,
          allocateBlockId: context.allocateBlockId,
        },
        {
          rowId: anchor!,
          where,
          cells,
          ...(lists.length ? { lists } : {}),
          ...(cellProperties.length ? { cellProperties } : {}),
          ...(codec.values(row).height === undefined
            ? {}
            : { properties: { heightTwips: codec.values(row).height } }),
        }
      );
      if (!result.success) return result;
      for (const edit of result.plan.operations) {
        if (edit.kind !== 'insert-row') return fail('Unexpected native row insertion intent.');
        rowCellIds(inserted.part, fresh).forEach((cellId, cellIndex) =>
          allocated.push({ cellId, operation: edit, cellIndex })
        );
        preview = inserted.part;
        operations.push(edit);
      }
    }
  }
  // Topology changes cannot silently transfer a retained cell's content to another merge owner.
  for (const row of pairs)
    for (const column of alignedColumns.alignment.pairs) {
      const owner = bindings.get([...beforePath, row.before, column.before].join('/'))?.table;
      if (owner?.cellId && !findNode(preview, owner.cellId))
        return fail('A changed merged cell owner needs native source reconciliation.');
    }
  const heads = new Map<number, string>();
  for (const cell of input.projection.cells
    .filter((cell) => cell.tableId === tableId)
    .sort((a, b) => a.rowIndex - b.rowIndex)) {
    if (!cell.verticalContinuation) heads.set(cell.gridColumn, cell.id);
    else if (
      findNode(preview, cell.id) &&
      heads.has(cell.gridColumn) &&
      !findNode(preview, heads.get(cell.gridColumn)!)
    )
      return fail('A promoted vertical merge owner needs native source reconciliation.');
  }
  let shapeGeometry: readonly CanonicalEdit[] = [];
  if (changedShape) {
    const shaped = planTargetShapeGeometry(
      context,
      codec,
      expansion.expanded,
      preview,
      operations,
      allocated
    );
    if (!shaped.success) return shaped;
    operations = [...shaped.topology];
    shapeGeometry = shaped.operations;
    for (const key of shaped.handled) handled.add(key);
  }
  for (const pair of pairs) {
    const old = oldRows[pair.before]!,
      next = newRows[pair.after]!;
    if (codec.key(old) !== codec.key(next))
      return fail('Changed row metadata needs explicit native property intentions.');
    for (const column of alignedColumns.alignment.pairs) {
      const oldCell = old.children[column.before] as LogicalGroup,
        newCell = next.children[column.after] as LogicalGroup,
        oldPath = [...beforePath, pair.before, column.before],
        newPath = [...afterPath, pair.after, column.after],
        owner = bindings.get(oldPath.join('/'))?.table;
      if (handled.has(`${pair.before}/${column.before}`)) continue;
      if (codec.key(oldCell) !== codec.key(newCell))
        return fail('Changed cell metadata needs explicit native property intentions.');
      if (!owner || owner.tableId !== tableId || owner.rowId !== rowIds[pair.before])
        return fail('A retained cell has no exact canonical table/row binding.');
      if (context.same(oldCell, newCell)) continue;
      if (sameTargetCellContent(context, oldCell.children, newCell.children, oldPath)) continue;
      if (owner.role !== 'cell' || !owner.cellId || !owner.rowId)
        return fail('Generated merged slots cannot become editable cell content.');
      const checked = planTableCellEdits(
        {
          revision,
          part,
          projection: input.projection,
          displayMode: input.projection.displayMode!,
          tableId,
          rowId: owner.rowId,
          cellId: owner.cellId,
        },
        []
      );
      if (!checked.success) return checked;
      const error = context.visitCell(
        oldCell.children,
        blankCellTarget(oldCell.children, newCell.children, newPath),
        oldPath,
        newPath,
        owner.cellId
      );
      if (error) return error;
    }
  }
  const formatting = planTargetTableProperties(
    context,
    codec,
    pairs,
    alignedColumns.alignment.pairs,
    preview,
    new Set(
      requests.flatMap((request) =>
        request.kind === 'split'
          ? request.parts.slice(1).map((part) => `${part.row}/${part.column}`)
          : []
      )
    )
  );
  if (!formatting.success) return formatting;
  const geometry = planTargetGeometry(context, codec, requests, grid, preview);
  if (!geometry.success) return geometry;
  return createEditPlan(revision, [
    ...operations,
    ...formatting.plan.operations,
    ...geometry.operations,
    ...shapeGeometry,
  ]);
}
