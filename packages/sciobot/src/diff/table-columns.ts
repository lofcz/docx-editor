import { applyTreeOp, findNode, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import {
  createEditPlan,
  EDIT_PLAN_LIMITS,
  type EditPlanResult,
  type PlannedTableOperation,
} from './plan';

export interface TableColumnSource {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  readonly tableId: string;
}
export type TableColumnTarget =
  | {
      readonly kind: 'insert';
      readonly where: 'left' | 'right';
      /** Required when the canonical table declares its grid. */
      readonly gridColumnId: string;
    }
  | {
      readonly kind: 'insert';
      readonly where: 'left' | 'right';
      /** An existing first-row cell, only for a native table without grid columns. */
      readonly referenceCellId: string;
    }
  | { readonly kind: 'delete'; readonly gridColumnId: string };

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const id = (value: unknown): value is string => typeof value === 'string' && value.length > 0;

/**
 * Plan one native column edit. Core owns grid widths, span changes, vertical merges, inheritance,
 * and bounded topology. The immutable preview never supplies allocated IDs to the plan.
 * A01/R09 own guarded execution and review. D17 owns rich new cell content and deferred IDs.
 */
export function planTableColumn(
  source: TableColumnSource,
  target: TableColumnTarget
): EditPlanResult {
  if (!source || typeof source !== 'object') return fail('Missing native table column source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'part', 'tableId'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Native column source fields must be plain values.');
  const input = { revision: source.revision, tableId: source.tableId, target };
  const problem = jsonDataProblem(input, EDIT_PLAN_LIMITS);
  if (problem) return fail(problem);
  const raw = JSON.stringify(input);
  if (
    raw.length > EDIT_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
  )
    return fail('Table column input exceeds its byte limit.');
  if (!target || typeof target !== 'object' || Array.isArray(target) || !id(source.tableId))
    return fail('Invalid native table column target.');
  const grid = Object.hasOwn(target, 'gridColumnId'),
    cell = Object.hasOwn(target, 'referenceCellId');
  if (
    (target.kind !== 'insert' && target.kind !== 'delete') ||
    Object.keys(target).some(
      (key) =>
        !(
          target.kind === 'insert'
            ? ['kind', 'where', 'gridColumnId', 'referenceCellId']
            : ['kind', 'gridColumnId']
        ).includes(key)
    ) ||
    (target.kind === 'insert' &&
      ((target.where !== 'left' && target.where !== 'right') || grid === cell)) ||
    (grid && !id((target as { gridColumnId?: unknown }).gridColumnId)) ||
    (cell && !id((target as { referenceCellId?: unknown }).referenceCellId)) ||
    (target.kind === 'delete' && !grid)
  )
    return fail('Invalid native table column anchor or side.');
  const table = findNode(source.part, source.tableId);
  if (table?.kind !== 'table') return fail('The native table does not exist.', 'invalid-document');
  const operation: PlannedTableOperation =
    target.kind === 'delete'
      ? { op: 'deleteTableColumn', tableId: table.id, gridColumnId: target.gridColumnId }
      : 'gridColumnId' in target
        ? {
            op: 'insertTableColumn',
            tableId: table.id,
            gridColumnId: target.gridColumnId,
            where: target.where,
          }
        : {
            op: 'insertTableColumn',
            tableId: table.id,
            referenceCellId: target.referenceCellId,
            where: target.where,
          };
  // Use the public dispatcher: it selects the regular or merged native column implementation.
  const preview = applyTreeOp(source.part, operation);
  if (!preview.ok)
    return fail(`Native column planning refused: ${preview.reason}.`, 'unsupported-operation');
  return createEditPlan(source.revision, [
    { kind: 'table', partName: source.part.name, operation },
  ]);
}
