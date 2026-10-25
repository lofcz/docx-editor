import { targetNewCellProperties, type NewCellPropertyBinding } from './new-cell-properties';
import type { TargetTablePropertyCodec } from './target-table-property-codec';
import type { InsertedCellListBinding } from './inserted-cell-lists';
import { rowCellIds } from './fragment-seeds';
import { applyTreeOp, withPart, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { LogicalGroup } from './logical-block-types';
import type { CanonicalEdit } from './plan';
import type { TargetTableContext } from './target-table';
import {
  alignTargetSequence,
  targetKeyCounts,
  type TargetSequenceAlignment,
} from './target-sequence';
import { nativeColumnAnchors } from './inserted-column-native';
import { planTableColumn } from './table-columns';
import { planInsertedTableColumn } from './inserted-table-columns';
import { insertedTargetCellFragments } from './target-table-fragments';
import { compileInlineTarget } from './inline-target';

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

/** Prefer complete vectors, header text, and unambiguous body evidence. Repeated hints retain position. */
export function alignTargetColumns(
  before: readonly LogicalGroup[],
  after: readonly LogicalGroup[],
  contentKey: (cell: LogicalGroup) => string = (cell) => cell.key
) {
  const vectors = (rows: readonly LogicalGroup[]) =>
      rows[0]!.children.map((_, column) =>
        JSON.stringify(rows.map((row) => contentKey(row.children[column] as LogicalGroup)))
      ),
    oldVectors = vectors(before),
    newVectors = vectors(after),
    oldCounts = targetKeyCounts(oldVectors),
    newCounts = targetKeyCounts(newVectors),
    oldKeys = oldVectors.map((key, index) =>
      oldCounts.get(key) === 1 && newCounts.get(key) === 1 ? `vector:${key}` : `old:${index}`
    ),
    newKeys = newVectors.map((key, index) =>
      newCounts.get(key) === 1 && oldCounts.get(key) === 1 ? `vector:${key}` : `new:${index}`
    ),
    headerKey = (cell: LogicalGroup) => {
      const child = cell.children[0];
      if (cell.children.length === 1 && child?.kind === 'paragraph') {
        const inline = compileInlineTarget(child.ast);
        if (inline.success && inline.target.plainText !== null)
          return `text:${inline.target.plainText}`;
      }
      return `cell:${contentKey(cell)}`;
    },
    headers = (rows: readonly LogicalGroup[], keys: readonly string[]) =>
      rows[0]!.children.map((cell, index) =>
        keys[index]!.startsWith('vector:') ? null : headerKey(cell as LogicalGroup)
      ),
    oldHeaders = headers(before, oldKeys),
    newHeaders = headers(after, newKeys),
    oldHeaderCounts = targetKeyCounts(oldHeaders.filter((key): key is string => key !== null)),
    newHeaderCounts = targetKeyCounts(newHeaders.filter((key): key is string => key !== null));
  for (const [keys, names] of [
    [oldKeys, oldHeaders],
    [newKeys, newHeaders],
  ] as const)
    names.forEach((key, index) => {
      if (key !== null && oldHeaderCounts.get(key) === 1 && newHeaderCounts.get(key) === 1)
        keys[index] = `header:${key}`;
    });
  const bodyOwners = (rows: readonly LogicalGroup[]) => {
    const owners = new Map<string, Set<number>>();
    for (const row of rows.slice(1))
      row.children.forEach((cell, column) => {
        const key = contentKey(cell as LogicalGroup),
          list = owners.get(key) ?? new Set<number>();
        list.add(column);
        owners.set(key, list);
      });
    return owners;
  };
  const oldBody = bodyOwners(before),
    newBody = bodyOwners(after),
    oldCandidates = new Map<number, Set<number>>(),
    newCandidates = new Map<number, Set<number>>();
  for (const [key, owners] of oldBody) {
    const next = newBody.get(key);
    if (owners.size !== 1 || next?.size !== 1) continue;
    const a = [...owners][0]!,
      b = [...next][0]!;
    if (!oldKeys[a]!.startsWith('old:') || !newKeys[b]!.startsWith('new:')) continue;
    const target = oldCandidates.get(a) ?? new Set<number>(),
      original = newCandidates.get(b) ?? new Set<number>();
    target.add(b);
    original.add(a);
    oldCandidates.set(a, target);
    newCandidates.set(b, original);
  }
  for (const [a, candidates] of oldCandidates) {
    if (candidates.size !== 1) continue;
    const b = [...candidates][0]!;
    if (newCandidates.get(b)?.size !== 1) continue;
    oldKeys[a] = newKeys[b] = `body:${a}:${b}`;
  }
  return alignTargetSequence(oldKeys, newKeys);
}

/** Materialize topology only in an immutable preview. Every returned address comes from the source. */
export function planTargetColumns(
  context: TargetTableContext,
  alignment: TargetSequenceAlignment,
  rows: TargetSequenceAlignment['pairs'],
  rowIds: readonly string[],
  codec: TargetTablePropertyCodec
):
  | {
      readonly success: true;
      readonly operations: readonly CanonicalEdit[];
      readonly part: OoxmlPart;
      readonly allocated: readonly {
        readonly cellId: string;
        readonly operation: CanonicalEdit;
        readonly cellIndex: number;
      }[];
    }
  | NativeEditFailure {
  const { input, source, after } = context,
    original = input.pkg.parts.get(source.index.partName)!,
    tableId = context.binding.table!.tableId,
    anchors = nativeColumnAnchors(original, tableId),
    operations: CanonicalEdit[] = [],
    allocated: { cellId: string; operation: CanonicalEdit; cellIndex: number }[] = [],
    targetRows = new Map(
      rows.map((pair) => [rowIds[pair.before]!, after.children[pair.after] as LogicalGroup])
    );
  let part = original;
  for (const gap of [...alignment.gaps].reverse()) {
    for (let index = gap.beforeEnd - 1; index >= gap.beforeStart; index--) {
      const planned = planTableColumn(
        { revision: source.index.revision, part, tableId },
        {
          kind: 'delete',
          ...(anchors[index]!.kind === 'grid-column'
            ? { gridColumnId: anchors[index]!.id }
            : { referenceCellId: anchors[index]!.id }),
        }
      );
      if (!planned.success) return planned;
      for (const edit of planned.plan.operations) {
        if (edit.kind !== 'table') return fail('Unexpected native column deletion intent.');
        const result = applyTreeOp(part, edit.operation);
        if (!result.ok) return fail(`Combined column deletion refused: ${result.reason}.`);
        part = result.part;
        operations.push(edit);
      }
    }
    const anchor = anchors[gap.beforeStart] ?? anchors[gap.beforeStart - 1],
      where = anchors[gap.beforeStart] ? 'left' : 'right',
      inserted = Array.from(
        { length: gap.afterEnd - gap.afterStart },
        (_, index) => gap.afterStart + index
      );
    if (where === 'right') inserted.reverse();
    for (const column of inserted) {
      const candidate = applyTreeOp(part, {
        op: 'insertTableColumn',
        tableId,
        ...(anchor!.kind === 'grid-column'
          ? { gridColumnId: anchor!.id }
          : { referenceCellId: anchor!.id }),
        where,
      });
      if (!candidate.ok) return fail(`Combined column insertion refused: ${candidate.reason}.`);
      const cells = [];
      const lists: InsertedCellListBinding[] = [];
      const cellProperties: NewCellPropertyBinding[] = [];
      for (const rowId of rowIds) {
        const row = targetRows.get(rowId);
        if (!row) {
          cells.push({ rowId, blocks: [] });
          continue;
        }
        const fragments = insertedTargetCellFragments(
          row.children[column] as LogicalGroup,
          context.targetAst,
          context.astOf,
          { revision: source.index.revision, pkg: input.pkg, partName: part.name },
          codec
        );
        if (!fragments.success) return fragments;
        const old = new Set(rowCellIds(part, rowId));
        const creates = rowCellIds(candidate.part, rowId).some((id) => !old.has(id));
        if (creates)
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
          row.children[column] as LogicalGroup,
          column
        );
        if (creates && Object.keys(properties).length)
          cellProperties.push({ cellIndex: cells.length, properties });
        else if (!creates && Object.keys(properties).length)
          return fail('A widened merged cell needs original property ownership.');
        cells.push({ rowId, blocks: creates ? fragments.fragments : [] });
      }
      const planned = planInsertedTableColumn(
        {
          revision: source.index.revision,
          pkg: withPart(input.pkg, part),
          partName: part.name,
          tableId,
          allocateBlockId: context.allocateBlockId,
        },
        {
          anchor: anchor!,
          where,
          cells,
          ...(lists.length ? { lists } : {}),
          ...(cellProperties.length ? { cellProperties } : {}),
        }
      );
      if (!planned.success) return planned;
      for (const edit of planned.plan.operations) {
        if (edit.kind !== 'insert-column')
          return fail('Unexpected native column insertion intent.');
        const result = candidate;
        edit.intent.cells.forEach((entry, cellIndex) => {
          const old = new Set(rowCellIds(part, entry.rowId));
          for (const cellId of rowCellIds(result.part, entry.rowId))
            if (!old.has(cellId)) allocated.push({ cellId, operation: edit, cellIndex });
        });
        part = result.part;
        operations.push(edit);
      }
    }
  }
  return { success: true, operations, part, allocated };
}
