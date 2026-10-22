import { targetNewCellProperties, type NewCellPropertyBinding } from './new-cell-properties';
import type { TargetAstNode } from '../projection/parse-target';
import { applyTreeOp, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { LogicalBlock, LogicalGroup } from './logical-block-types';
import type { CanonicalEdit } from './plan';
import type { TargetTableContext } from './target-table';
import type { TargetTablePropertyCodec } from './target-table-property-codec';
import { blankCellTarget, sameTargetCellContent } from './target-table-content';
import { insertedTargetCellFragments } from './target-table-fragments';
import {
  nativeTableSlotGrid,
  resolveTableGeometry,
  type TableGeometryIntent,
} from './table-geometry';
import {
  targetGeometryRequests,
  sourceSlotGrid,
  targetSlotKey as slot,
  type ExpandedTargetTable,
  type TargetSlotRegion,
} from './target-table-spans';

export interface AllocatedTargetCell {
  readonly cellId: string;
  readonly operation: CanonicalEdit;
  readonly cellIndex: number;
}
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

/** Rebase geometry onto topology without lending allocated preview identities to an edit plan. */
export function planTargetShapeGeometry(
  context: TargetTableContext,
  codec: TargetTablePropertyCodec,
  target: ExpandedTargetTable,
  initial: OoxmlPart,
  topology: readonly CanonicalEdit[],
  allocated: readonly AllocatedTargetCell[]
):
  | {
      readonly success: true;
      readonly operations: readonly CanonicalEdit[];
      readonly topology: readonly CanonicalEdit[];
      readonly handled: ReadonlySet<string>;
    }
  | NativeEditFailure {
  const tableId = context.binding.table!.tableId,
    grid = nativeTableSlotGrid(initial, tableId),
    targetRows = context.after.children as readonly LogicalGroup[],
    width = targetRows[0]!.children.length;
  if (
    !grid ||
    grid.heads.length !== targetRows.length ||
    grid.heads.some((row) => row.length !== width)
  )
    return fail('The planned topology has no exact target slot grid.');
  const compared = targetGeometryRequests(grid, target);
  if (!compared.success) return compared;
  const oldCells = new Map<string, { cell: LogicalGroup; path: readonly number[] }>();
  (context.before.children as readonly LogicalGroup[]).forEach((row, r) =>
    row.children.forEach((cell, c) => {
      const path = [...context.beforePath, r, c],
        owner = context.bindings.get(path.join('/'))?.table;
      if (owner?.role === 'cell' && owner.cellId)
        oldCells.set(owner.cellId, { cell: cell as LogicalGroup, path });
    })
  );
  const born = new Map(allocated.map((entry) => [entry.cellId, entry])),
    changedSeeds = new Map<CanonicalEdit, Map<number, readonly LogicalBlock[]>>(),
    handled = new Set<string>(),
    operations: CanonicalEdit[] = [];
  const originalGrid = sourceSlotGrid(
    context.bindings,
    context.beforePath,
    context.before.children.length,
    (context.before.children[0] as LogicalGroup).children.length
  );
  const markHandled = (path: readonly number[]) => {
    const [row, column] = path.slice(context.beforePath.length),
      region = originalGrid?.regions.get(slot(row!, column!));
    if (!region) {
      handled.add(path.slice(context.beforePath.length).join('/'));
      return;
    }
    for (let r = region.row; r < region.row + region.rows; r++)
      for (let c = region.column; c < region.column + region.columns; c++) handled.add(slot(r, c));
  };
  const sourceAt = (head: TargetSlotRegion) => {
    const id = grid.cellIds.get(slot(head.row, head.column))!,
      old = oldCells.get(id);
    return {
      id,
      old,
      cell: old?.cell ?? (targetRows[head.row]!.children[head.column] as LogicalGroup),
      path: old?.path ?? [...context.afterPath, head.row, head.column],
    };
  };
  const visit = (
    head: TargetSlotRegion,
    blocks: readonly LogicalBlock[],
    targetHead: TargetSlotRegion
  ): NativeEditFailure | null => {
    const source = sourceAt(head);
    if (source.old) {
      markHandled(source.path);
      if (sameTargetCellContent(context, source.cell.children, blocks, source.path)) return null;
      return context.visitCell(
        source.cell.children,
        blankCellTarget(source.cell.children, blocks, [
          ...context.afterPath,
          targetHead.row,
          targetHead.column,
        ]),
        source.path,
        [...context.afterPath, targetHead.row, targetHead.column],
        source.id
      );
    }
    const allocation = born.get(source.id);
    if (!allocation)
      return fail('A changed table cell has neither a retained owner nor an allocation intention.');
    const changes =
      changedSeeds.get(allocation.operation) ?? new Map<number, readonly LogicalBlock[]>();
    changes.set(allocation.cellIndex, blocks);
    changedSeeds.set(allocation.operation, changes);
    return null;
  };
  let part = initial;
  for (const request of compared.requests) {
    let intent: TableGeometryIntent;
    if (request.kind === 'merge') {
      const anchor = targetRows[request.region.row]!.children[
        request.region.column
      ] as LogicalGroup;
      if (codec.key(anchor) !== codec.key(sourceAt(request.heads[0]!).cell))
        return fail('Changed merged cell metadata needs native property destinations.');
      let end = anchor.children.length,
        retaining = true;
      for (let i = request.heads.length - 1; i >= 1; i--) {
        const head = request.heads[i]!,
          source = sourceAt(head),
          blocks = source.cell.children;
        if (source.old) markHandled(source.path);
        if (!blocks.length) continue;
        if (
          retaining &&
          blocks.length <= end &&
          sameTargetCellContent(
            context,
            blocks,
            anchor.children.slice(end - blocks.length, end),
            source.path
          )
        ) {
          end -= blocks.length;
          continue;
        }
        retaining = false;
        const error = visit(head, [], request.region);
        if (error) return error;
      }
      const error = visit(request.heads[0]!, anchor.children.slice(0, end), request.region);
      if (error) return error;
      intent = { tableId, kind: 'merge', region: request.region };
    } else {
      const first = request.parts[0]!,
        next = targetRows[first.row]!.children[first.column] as LogicalGroup;
      const error = visit(request.region, next.children, first);
      if (error) return error;
      const cells = [];
      const cellProperties: NewCellPropertyBinding[] = [];
      for (const region of request.parts.slice(1)) {
        const cell = targetRows[region.row]!.children[region.column] as LogicalGroup,
          fragments = insertedTargetCellFragments(
            cell,
            context.targetAst,
            context.astOf,
            undefined,
            codec
          );
        const properties = targetNewCellProperties(
          codec,
          context.after,
          targetRows[region.row]!,
          cell,
          region.column
        );
        if (Object.keys(properties).length)
          cellProperties.push({ cellIndex: cells.length, properties });
        if (!fragments.success) return fragments;
        cells.push(fragments.fragments);
      }
      intent = {
        tableId,
        kind: 'split',
        region: request.region,
        rows: request.rows,
        columns: request.columns,
        cells,
        ...(cellProperties.length ? { cellProperties } : {}),
      };
    }
    const resolved = resolveTableGeometry(part, intent);
    if (!resolved.success) return resolved;
    const result = applyTreeOp(part, resolved.operation);
    if (!result.ok) return fail(`Native compound geometry refused: ${result.reason}.`);
    part = result.part;
    operations.push({ kind: 'table-geometry', partName: part.name, intent });
  }
  const revised: CanonicalEdit[] = [];
  for (const edit of topology) {
    const changes = changedSeeds.get(edit);
    if (!changes) {
      revised.push(edit);
      continue;
    }
    if (edit.kind !== 'insert-row' && edit.kind !== 'insert-column')
      return fail('A born cell has no native topology seed.');
    const streams: (readonly TargetAstNode[])[] = [];
    for (let index = 0; index < edit.intent.cells.length; index++) {
      const blocks = changes.get(index);
      if (!blocks) {
        streams.push(
          edit.kind === 'insert-row' ? edit.intent.cells[index]! : edit.intent.cells[index]!.blocks
        );
        continue;
      }
      const fragments = insertedTargetCellFragments(
        {
          kind: 'cell',
          header: { type: 'cell' },
          domainType: null,
          children: blocks,
          path: [],
          list: null,
          key: 'native-seed',
          rowGroups: [],
        },
        context.targetAst,
        context.astOf
      );
      if (!fragments.success) return fragments;
      streams.push(fragments.fragments);
    }
    if (edit.kind === 'insert-row')
      revised.push({
        ...edit,
        intent: {
          ...edit.intent,
          cells: streams,
        },
      });
    else
      revised.push({
        ...edit,
        intent: {
          ...edit.intent,
          cells: edit.intent.cells.map((cell, index) => ({
            ...cell,
            blocks: streams[index]!,
          })),
        },
      });
  }
  return { success: true, operations, topology: revised, handled };
}
