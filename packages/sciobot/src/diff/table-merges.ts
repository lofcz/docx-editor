import { applyTreeOp, findNode, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import {
  createEditPlan,
  EDIT_PLAN_LIMITS,
  type EditPlanResult,
  type PlannedTableOperation,
} from './plan';

export interface TableMergeSource {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  readonly tableId: string;
}
export type TableMergeTarget =
  | { readonly kind: 'merge'; readonly cellIds: readonly string[] }
  | {
      readonly kind: 'split';
      /** Any canonical part of the addressed visual cell, including a vertical continuation. */
      readonly cellId: string;
      readonly rows: number;
      readonly columns: number;
      /** Existing-grid restores authored slots; omission uses native equal-width splitting. */
      readonly columnDistribution?: 'equal-width' | 'existing-grid';
    };
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const id = (value: unknown): value is string => typeof value === 'string' && value.length > 0;
const count = (value: unknown): value is number => Number.isSafeInteger(value) && Number(value) > 0;

/**
 * Plan one native merge/split without rebuilding cells or copying content. Core owns rectangles,
 * visual merge chains, grid cuts, restored slots, content movement, locks, and resource limits.
 * Allocated preview IDs stay outside plans. D15 composes targets; A01/R10 own guarded execution.
 */
export function planTableMerge(source: TableMergeSource, target: TableMergeTarget): EditPlanResult {
  if (!source || typeof source !== 'object') return fail('Missing native merge/split source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'part', 'tableId'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Native merge/split source fields must be plain values.');
  const input = { revision: source.revision, tableId: source.tableId, target };
  const problem = jsonDataProblem(input, EDIT_PLAN_LIMITS);
  if (problem) return fail(problem);
  const raw = JSON.stringify(input);
  if (
    raw.length > EDIT_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
  )
    return fail('Native merge/split input exceeds its byte limit.');
  if (
    !target ||
    typeof target !== 'object' ||
    Array.isArray(target) ||
    !id(source.tableId) ||
    (target.kind !== 'merge' && target.kind !== 'split') ||
    Object.keys(target).some(
      (key) =>
        !(
          target.kind === 'merge'
            ? ['kind', 'cellIds']
            : ['kind', 'cellId', 'rows', 'columns', 'columnDistribution']
        ).includes(key)
    )
  )
    return fail('Invalid native merge/split target.');
  if (target.kind === 'merge') {
    if (
      !Array.isArray(target.cellIds) ||
      !target.cellIds.length ||
      !target.cellIds.every(id) ||
      new Set(target.cellIds).size !== target.cellIds.length
    )
      return fail('Merge targets need unique canonical cell IDs.');
  } else if (
    !id(target.cellId) ||
    !count(target.rows) ||
    !count(target.columns) ||
    (Object.hasOwn(target, 'columnDistribution') &&
      target.columnDistribution !== 'equal-width' &&
      target.columnDistribution !== 'existing-grid')
  )
    return fail('Invalid native split counts or column distribution.');
  const table = findNode(source.part, source.tableId);
  if (table?.kind !== 'table') return fail('The native table does not exist.', 'invalid-document');
  const operation: PlannedTableOperation =
    target.kind === 'merge'
      ? { op: 'mergeTableCells', tableId: table.id, cellIds: target.cellIds }
      : {
          op: 'splitTableCell',
          tableId: table.id,
          cellId: target.cellId,
          rows: target.rows,
          columns: target.columns,
          ...(target.columnDistribution === undefined
            ? {}
            : { columnDistribution: target.columnDistribution }),
        };
  const preview = applyTreeOp(source.part, operation);
  if (!preview.ok)
    return fail(`Native merge/split planning refused: ${preview.reason}.`, 'unsupported-operation');
  return createEditPlan(source.revision, [
    { kind: 'table', partName: source.part.name, operation },
  ]);
}
