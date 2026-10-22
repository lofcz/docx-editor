import {
  applyTreeOp,
  findNode,
  type OoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import { checkNativeCellScope } from './cell-scope';
import { insertedColumnOperation, insertedColumnWidthOperation } from './inserted-column-native';
import {
  isInsertedColumnIntent,
  type InsertedTableColumnIntent,
  type InsertedColumnProperties,
  type InsertedColumnPayloadBinding,
} from './inserted-column-types';
import { insertedPayloadProblem } from './inserted-payloads';
import { createEditPlan, EDIT_PLAN_LIMITS, type EditPlanResult } from './plan';
import { planTableColumn } from './table-columns';

export type {
  InsertedTableColumnIntent,
  InsertedColumnAnchor,
  InsertedColumnCell,
  InsertedColumnProperties,
  InsertedColumnPayloadBinding,
} from './inserted-column-types';
export interface InsertedTableColumnSource {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  readonly partName: string;
  readonly tableId: string;
}
export type InsertedTableColumnTarget = Omit<
  InsertedTableColumnIntent,
  'tableId' | 'properties' | 'payloads'
> & {
  readonly properties?: InsertedColumnProperties;
  readonly payloads?: readonly InsertedColumnPayloadBinding[];
};
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;

/** Plan rich new column cells by original row IDs, with no live writes or preview identities. */
export function planInsertedTableColumn(
  source: InsertedTableColumnSource,
  target: InsertedTableColumnTarget
): EditPlanResult {
  if (!source || typeof source !== 'object') return fail('Missing inserted table column source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'pkg', 'partName', 'tableId'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Inserted column source fields must be plain values.');
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
    return fail('Inserted column input exceeds its byte limit.');
  if (
    !target ||
    typeof target !== 'object' ||
    Array.isArray(target) ||
    Object.keys(target).some(
      (key) => !['anchor', 'where', 'cells', 'properties', 'payloads', 'lists'].includes(key)
    )
  )
    return fail('Invalid inserted column target.');
  const intent = {
    tableId: source.tableId,
    ...target,
    properties: target.properties === undefined ? {} : target.properties,
    payloads: target.payloads === undefined ? [] : target.payloads,
  };
  if (!isInsertedColumnIntent(intent))
    return fail('Invalid inserted column fragments, properties, or payload bindings.');
  const part = source.pkg.parts.get(source.partName);
  if (!part) return fail('The canonical table part does not exist.', 'invalid-document');
  const base = planTableColumn(
    { revision: source.revision, part, tableId: source.tableId },
    {
      kind: 'insert',
      where: intent.where,
      ...(intent.anchor.kind === 'grid-column'
        ? { gridColumnId: intent.anchor.id }
        : { referenceCellId: intent.anchor.id }),
    }
  );
  if (!base.success) return base;
  const preview = applyTreeOp(part, insertedColumnOperation(intent));
  if (!preview.ok)
    return fail(`Native column planning refused: ${preview.reason}.`, 'unsupported-operation');
  const rows = children(findNode(part, intent.tableId)!).filter((node) => node.kind === 'tableRow');
  if (rows.length !== intent.cells.length)
    return fail('Column streams must address every original native row exactly once.');
  const byRow = new Map(intent.cells.map((cell) => [cell.rowId, cell.blocks]));
  for (const row of rows) {
    const blocks = byRow.get(row.id);
    if (!blocks) return fail('Column streams contain a foreign or missing native row.');
    const old = new Set(
      children(row)
        .filter((node) => node.kind === 'tableCell')
        .map((node) => node.id)
    );
    const created = children(findNode(preview.part, row.id)!).filter(
      (node) => node.kind === 'tableCell' && !old.has(node.id)
    );
    if (created.length > 1)
      return fail('Native column insertion created ambiguous cells.', 'invalid-document');
    if (!created.length) {
      if (blocks.length)
        return fail('This row widens an existing span or ends before the inserted column.');
      continue;
    }
    const scope = checkNativeCellScope(preview.part, {
      tableId: intent.tableId,
      rowId: row.id,
      cellId: created[0]!.id,
    });
    if (scope && blocks.length) return scope;
  }
  const sizing = insertedColumnWidthOperation(part, preview.part, intent);
  if (!sizing.ok) return fail(sizing.reason, 'unsupported-operation');
  if (sizing.operation) {
    const sized = applyTreeOp(preview.part, sizing.operation);
    if (!sized.ok)
      return fail(
        `Native column width planning refused: ${sized.reason}.`,
        'unsupported-operation'
      );
  }
  const payloadProblem = insertedPayloadProblem(source.pkg, intent.payloads);
  if (payloadProblem) return payloadProblem;
  return createEditPlan(source.revision, [{ kind: 'insert-column', partName: part.name, intent }]);
}
