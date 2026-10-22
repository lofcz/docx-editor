import { applyTreeOp, findNode, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import { nearestCellAncestor } from './cell-scope';
import {
  createEditPlan,
  EDIT_PLAN_LIMITS,
  type EditPlanResult,
  type PlannedTableOperation,
} from './plan';

export interface TableRowSource {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  readonly tableId: string;
}
export type TableRowTarget =
  | {
      readonly kind: 'insert';
      readonly rowId: string;
      readonly where: 'above' | 'below';
      /** Complete values for one new ordinary row. Omission inserts the native empty row skeleton. */
      readonly values?: readonly string[];
    }
  | {
      readonly kind: 'delete';
      readonly rowId: string;
      readonly referenceCellId?: string;
    };
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/**
 * Plan one native row operation, including inherited row/cell properties and optional plain values.
 * A detached native preview validates geometry and locks. Allocated preview IDs never enter the plan.
 * D16 owns rich inserted fragments and explicit property overrides. A01/R08 own guarded execution.
 */
export function planTableRow(source: TableRowSource, target: TableRowTarget): EditPlanResult {
  if (!source || typeof source !== 'object') return fail('Missing native table row source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'part', 'tableId'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Native table row source fields must be plain values.');
  const input = { revision: source.revision, tableId: source.tableId, target };
  const problem = jsonDataProblem(input, EDIT_PLAN_LIMITS);
  if (problem) return fail(problem);
  const raw = JSON.stringify(input);
  if (
    raw.length > EDIT_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
  )
    return fail('Table row input exceeds its byte limit.');
  if (
    !target ||
    typeof target !== 'object' ||
    Array.isArray(target) ||
    typeof source.tableId !== 'string' ||
    !source.tableId ||
    typeof target.rowId !== 'string' ||
    !target.rowId ||
    !['insert', 'delete'].includes(target.kind) ||
    Object.keys(target).some(
      (key) =>
        !(
          target.kind === 'insert'
            ? ['kind', 'rowId', 'where', 'values']
            : ['kind', 'rowId', 'referenceCellId']
        ).includes(key)
    ) ||
    (target.kind === 'insert' &&
      (!['above', 'below'].includes(target.where) ||
        (target.values !== undefined &&
          (!Array.isArray(target.values) ||
            target.values.some((value) => typeof value !== 'string'))))) ||
    (target.kind === 'delete' &&
      target.referenceCellId !== undefined &&
      (typeof target.referenceCellId !== 'string' || !target.referenceCellId))
  )
    return fail('Invalid native table row target.');
  const table = findNode(source.part, source.tableId),
    row = findNode(source.part, target.rowId);
  if (
    table?.kind !== 'table' ||
    row?.kind !== 'tableRow' ||
    nearestCellAncestor(source.part, row.id, 'table')?.id !== table.id
  )
    return fail('The row does not belong to the requested native table.', 'invalid-document');
  if (target.kind === 'delete' && target.referenceCellId !== undefined) {
    const cell = findNode(source.part, target.referenceCellId);
    if (
      cell?.kind !== 'tableCell' ||
      nearestCellAncestor(source.part, cell.id, 'tableRow')?.id !== row.id ||
      nearestCellAncestor(source.part, cell.id, 'table')?.id !== table.id
    )
      return fail('The reference cell does not belong to the removed row.', 'invalid-document');
  }
  const operation: PlannedTableOperation =
    target.kind === 'delete'
      ? {
          op: 'deleteTableRow',
          tableId: table.id,
          rowId: row.id,
          ...(target.referenceCellId === undefined
            ? {}
            : { referenceCellId: target.referenceCellId }),
        }
      : target.values === undefined
        ? { op: 'insertTableRow', tableId: table.id, rowId: row.id, where: target.where }
        : {
            op: 'authorTable',
            action: {
              kind: 'existing',
              tableId: table.id,
              mutation: {
                kind: 'insertRows',
                rowId: row.id,
                location: target.where === 'above' ? 'before' : 'after',
                count: 1,
                values: [target.values],
              },
            },
          };
  const preview = applyTreeOp(source.part, operation);
  if (!preview.ok)
    return fail(`Native row planning refused: ${preview.reason}.`, 'unsupported-operation');
  return createEditPlan(source.revision, [
    { kind: 'table', partName: source.part.name, operation },
  ]);
}
