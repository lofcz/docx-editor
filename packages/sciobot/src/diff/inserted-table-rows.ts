import type { NewCellPropertyBinding } from './new-cell-properties';
import type { InsertedCellListBinding } from './inserted-cell-lists';
import {
  applyTreeOp,
  findNode,
  type OoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import type { TargetAstNode } from '../projection/parse-target';
import { checkNativeCellScope } from './cell-scope';
import { insertedPayloadProblem } from './inserted-payloads';
import {
  isInsertedRowIntent,
  type InsertedRowPayloadBinding,
  type InsertedRowProperties,
} from './inserted-row-types';
import { createEditPlan, EDIT_PLAN_LIMITS, type EditPlanResult } from './plan';
import { planTableRow } from './table-rows';

export type {
  InsertedTableRowIntent,
  InsertedRowPayloadBinding,
  InsertedRowProperties,
} from './inserted-row-types';

export interface InsertedTableRowSource {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  readonly partName: string;
  readonly tableId: string;
}
export interface InsertedTableRowTarget {
  readonly rowId: string;
  readonly where: 'above' | 'below';
  readonly cells: readonly (readonly TargetAstNode[])[];
  readonly properties?: InsertedRowProperties;
  readonly payloads?: readonly InsertedRowPayloadBinding[];
  readonly lists?: readonly InsertedCellListBinding[];
  readonly cellProperties?: readonly NewCellPropertyBinding[];
}
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;

/**
 * Plan rich content for one new native row. The immutable native preview validates insertion,
 * physical cell count, inherited merges, explicit height, locks, and payload storage.
 * Only original anchors and new fragments survive. A01/R08 own guarded direct/review execution.
 */
export function planInsertedTableRow(
  source: InsertedTableRowSource,
  target: InsertedTableRowTarget
): EditPlanResult {
  if (!source || typeof source !== 'object') return fail('Missing inserted table row source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'pkg', 'partName', 'tableId'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Inserted row source fields must be plain values.');
  const input = {
    revision: source.revision,
    partName: source.partName,
    tableId: source.tableId,
    target,
  };
  const problem = jsonDataProblem(input, EDIT_PLAN_LIMITS);
  if (problem) return fail(problem);
  const raw = JSON.stringify(input);
  if (
    raw.length > EDIT_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
  )
    return fail('Inserted row input exceeds its byte limit.');
  if (
    !target ||
    typeof target !== 'object' ||
    Array.isArray(target) ||
    Object.keys(target).some(
      (key) =>
        !['rowId', 'where', 'cells', 'properties', 'payloads', 'lists', 'cellProperties'].includes(
          key
        )
    )
  )
    return fail('Invalid inserted row target.');
  const intent = {
    tableId: source.tableId,
    ...target,
    properties: target.properties === undefined ? {} : target.properties,
    payloads: target.payloads === undefined ? [] : target.payloads,
  };
  if (!isInsertedRowIntent(intent))
    return fail('Invalid inserted row fragments, properties, or payload bindings.');
  const part = source.pkg.parts.get(source.partName);
  if (!part) return fail('The canonical table part does not exist.', 'invalid-document');
  const base = planTableRow(
    { revision: source.revision, part, tableId: source.tableId },
    {
      kind: 'insert',
      rowId: target.rowId,
      where: target.where,
    }
  );
  if (!base.success) return base;
  const preview = applyTreeOp(part, {
    op: 'insertTableRow',
    tableId: source.tableId,
    rowId: target.rowId,
    where: target.where,
  });
  if (!preview.ok)
    return fail(`Native row planning refused: ${preview.reason}.`, 'unsupported-operation');
  const before = new Set(children(findNode(part, source.tableId)!).map((node) => node.id));
  const created = children(findNode(preview.part, source.tableId)!).filter(
    (node) => node.kind === 'tableRow' && !before.has(node.id)
  );
  if (created.length !== 1)
    return fail('Native insertion did not create exactly one row.', 'invalid-document');
  const row = created[0]!,
    cells = children(row).filter((node) => node.kind === 'tableCell');
  if (cells.length !== intent.cells.length)
    return fail('Rich row cells must match the native physical cell count.');
  for (let index = 0; index < cells.length; index++) {
    const scope = checkNativeCellScope(preview.part, {
      tableId: source.tableId,
      rowId: row.id,
      cellId: cells[index]!.id,
    });
    if (scope && intent.cells[index]!.length) return scope;
  }
  if (intent.properties.heightTwips !== undefined) {
    const height = applyTreeOp(preview.part, {
      op: 'setTableRowHeight',
      tableId: source.tableId,
      rowId: row.id,
      heightTwips: intent.properties.heightTwips,
    });
    if (!height.ok)
      return fail(`Native row height planning refused: ${height.reason}.`, 'unsupported-operation');
  }
  const payloadProblem = insertedPayloadProblem(source.pkg, intent.payloads);
  if (payloadProblem) return payloadProblem;
  return createEditPlan(source.revision, [{ kind: 'insert-row', partName: part.name, intent }]);
}
