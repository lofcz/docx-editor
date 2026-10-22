import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import {
  createEditPlan,
  EDIT_PLAN_LIMITS,
  type EditPlanResult,
  type PlannedTableOperation,
} from './plan';

const PROPERTY_KEYS = {
  setTableProperties: ['styleId', 'headerRowCount', 'columnWidthsTwips'],
  setTableColumnWidths: [
    'leftGridColumnId',
    'rightGridColumnId',
    'leftWidthTwips',
    'rightWidthTwips',
  ],
  setTableRightEdgeWidth: ['gridColumnId', 'columnWidthTwips', 'tableWidthTwips'],
  setTableRowHeight: ['rowId', 'heightTwips'],
  setTableCellBorders: ['cellIds', 'scope', 'target', 'spec'],
  setTableCellFill: ['cellIds', 'color'],
  setTableCellVerticalAlignment: ['cellIds', 'alignment'],
} as const satisfies Partial<Record<PlannedTableOperation['op'], readonly string[]>>;

export type PlannedTablePropertyOperation = Extract<
  PlannedTableOperation,
  { op: keyof typeof PROPERTY_KEYS }
>;
type WithoutTableId<T> = T extends unknown ? Omit<T, 'tableId'> : never;
/** Existing native property vocabulary. The source supplies the canonical table ID. */
export type TableFormatTarget = WithoutTableId<PlannedTablePropertyOperation>;
export interface TableFormatSource {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  readonly tableId: string;
}
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/**
 * Plan one native property command independently from cell content. Core validates units,
 * grid/merge geometry, selection completeness, themes, protection, and exact authored semantics.
 * A detached preview removes accepted authored no-ops. Unchanged native drag proposals refuse.
 * A01/R11 own guarded direct/review execution.
 */
export function planTableFormat(
  source: TableFormatSource,
  target: TableFormatTarget
): EditPlanResult {
  if (!source || typeof source !== 'object') return fail('Missing native table property source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'part', 'tableId'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Native table property source fields must be plain values.');
  const input = { revision: source.revision, tableId: source.tableId, target };
  const problem = jsonDataProblem(input, EDIT_PLAN_LIMITS);
  if (problem) return fail(problem);
  const raw = JSON.stringify(input);
  if (
    raw.length > EDIT_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
  )
    return fail('Table property input exceeds its byte limit.');
  if (
    !target ||
    typeof target !== 'object' ||
    Array.isArray(target) ||
    typeof source.tableId !== 'string' ||
    !source.tableId ||
    typeof target.op !== 'string' ||
    !Object.hasOwn(PROPERTY_KEYS, target.op) ||
    Object.keys(target).some(
      (key) => key !== 'op' && !(PROPERTY_KEYS[target.op] as readonly string[]).includes(key)
    )
  )
    return fail('Invalid native table property target.');
  const table = findNode(source.part, source.tableId);
  if (table?.kind !== 'table') return fail('The native table does not exist.', 'invalid-document');
  const operation: PlannedTablePropertyOperation = { ...target, tableId: table.id };
  const preview = applyTreeOp(source.part, operation);
  if (!preview.ok)
    return fail(
      `Native table property planning refused: ${preview.reason}.`,
      'unsupported-operation'
    );
  const updated = findNode(preview.part, table.id);
  if (updated?.kind !== 'table')
    return fail('The native property preview lost its table.', 'invalid-document');
  if (table === updated || canonicalOoxmlFingerprint(table) === canonicalOoxmlFingerprint(updated))
    return createEditPlan(source.revision, []);
  return createEditPlan(source.revision, [
    { kind: 'table', partName: source.part.name, operation },
  ]);
}
