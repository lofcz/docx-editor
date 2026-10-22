import { createDirectListContext } from '../ai/direct-lists';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  collectRevisionSites,
  findNode,
  parentNodeOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
  type RevisionAddress,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { CanonicalEditPlan } from '../diff/plan';
import {
  insertedRowSeeds,
  materializeInsertedRow,
  type InsertedRowSeeds,
} from './inserted-row-content';
import {
  nativePropertyAddress,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import type { TrackTextContext } from './tracked-inline';
import {
  commitTrackedStructure,
  fail,
  nodeIdsOf,
  preflightTrackedStructure,
  readTrackedStructureInput,
} from './tracked-structure';

/** Word row addresses survive save/reopen. Row IDs address this live editor revision only. */
export interface NativeRowProposal {
  readonly identity: string;
  readonly partName: string;
  readonly kind: 'rowInsert' | 'rowDelete';
  readonly tableId: string;
  readonly rowIds: readonly string[];
  readonly addresses: readonly RevisionAddress[];
}
export type TrackTableRowsResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposals: readonly NativeRowProposal[];
    }
  | NativeEditFailure;

interface RowStep {
  readonly op: TreeDocOp;
  readonly tableId: string;
  readonly kind: 'rowInsert' | 'rowDelete';
  /** Insertions: existing anchor and side. Deletions: the retained row. */
  readonly rowId: string;
  readonly where?: 'above' | 'below';
  readonly count: number;
  readonly seeds?: InsertedRowSeeds;
}

const rows = (part: OoxmlPart, tableId: string): readonly OoxmlNode[] => {
  const table = findNode(part, tableId);
  return table && table.kind !== 'textValue'
    ? table.children.filter((node) => node.kind === 'tableRow')
    : [];
};
const ownerRow = (part: OoxmlPart, node: OoxmlNode): OoxmlNode | null => {
  let current: OoxmlNode | null = node;
  for (let depth = 0; current && depth < 4; depth++) {
    if (current.kind === 'tableRow') return current;
    current = parentNodeOf(part, current.id);
  }
  return null;
};
const pendingRow = (row: OoxmlNode) =>
  row.kind !== 'textValue' &&
  row.children.some(
    (child) =>
      child.kind !== 'textValue' &&
      child.localName === 'trPr' &&
      child.children.some(
        (marker) =>
          marker.kind !== 'textValue' &&
          marker.namespaceUri === WML_NAMESPACE_URI &&
          ['ins', 'del', 'trPrChange'].includes(marker.localName)
      )
  );

function rowStep(
  edit: CanonicalEditPlan['operations'][number],
  revision: { author: string; date: string }
): RowStep | NativeEditFailure {
  if (edit.kind === 'insert-row') {
    const seeds = insertedRowSeeds(edit.intent);
    if (!seeds.success) return seeds;
    const { tableId, rowId, where } = edit.intent;
    return {
      op: { op: 'insertTableRow', tableId, rowId, where, revision },
      tableId,
      kind: 'rowInsert',
      rowId,
      where,
      count: 1,
      seeds: seeds.seeds,
    };
  }
  if (edit.kind !== 'table')
    return fail('This executor accepts table row plans only.', 'unsupported-operation');
  const operation = edit.operation;
  if (operation.op === 'insertTableRow')
    return {
      op: { ...operation, revision },
      tableId: operation.tableId,
      kind: 'rowInsert',
      rowId: operation.rowId,
      where: operation.where,
      count: 1,
    };
  if (operation.op === 'deleteTableRow')
    return {
      op: { ...operation, revision },
      tableId: operation.tableId,
      kind: 'rowDelete',
      rowId: operation.rowId,
      count: 1,
    };
  if (
    operation.op === 'authorTable' &&
    operation.action.kind === 'existing' &&
    operation.action.mutation.kind === 'insertRows'
  ) {
    const mutation = operation.action.mutation;
    return {
      op: { ...operation, revision },
      tableId: operation.action.tableId,
      kind: 'rowInsert',
      rowId: mutation.rowId,
      where: mutation.location === 'before' ? 'above' : 'below',
      count: mutation.count,
    };
  }
  return fail('This executor accepts table row plans only.', 'unsupported-operation');
}

/**
 * Propose native row insertions and deletions as Word row revisions in one gated,
 * separately undoable transaction. New row content belongs to its row insertion revision.
 */
export function trackTableRowPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext
): TrackTableRowsResult {
  const read = readTrackedStructureInput(editor, input, context, 'row');
  if (!read.success) return read;
  const { plan, actual, part, revision } = read.input;
  const steps: RowStep[] = [];
  const anchors = new Set<string>(),
    deleted = new Set<string>();
  for (const edit of plan.operations) {
    if (edit.partName !== part.name)
      return fail('Tracked rows must belong to the body part.', 'unsupported-operation');
    const step = rowStep(edit, revision);
    if ('success' in step) return step;
    const row = findNode(part, step.rowId);
    if (row?.kind !== 'tableRow' || !rows(part, step.tableId).includes(row))
      return fail('A planned row does not belong to its native table.', 'ambiguous-target');
    const protection = propertyControlRefusal(part, step.tableId);
    if (protection) return fail(protection, 'unsupported-operation');
    if (step.kind === 'rowDelete') {
      if (deleted.has(step.rowId)) return fail('A row is deleted twice.', 'ambiguous-target');
      if (pendingRow(row))
        return fail(
          'Changing an earlier pending row proposal needs explicit ownership reconciliation.',
          'unsupported-operation'
        );
      deleted.add(step.rowId);
    } else {
      const key = JSON.stringify([step.rowId, step.where]);
      if (anchors.has(key))
        return fail('Two row insertions share one native anchor.', 'ambiguous-target');
      anchors.add(key);
    }
    steps.push(step);
  }
  if (!steps.length)
    return { success: true, revision: Object.freeze(actual), proposals: Object.freeze([]) };
  const originalIds = nodeIdsOf(part);
  const originalSites = new Map(
    collectRevisionSites(part).map((site) => [site.node.id, JSON.stringify(site.node)])
  );
  let proposals: readonly NativeRowProposal[] = [];
  /** Locate rows this transaction created, fill them, and collect native row receipts. */
  const complete = (pkg: OoxmlPackage): OoxmlPackage => {
    const listContext = createDirectListContext();
    let working = pkg;
    const insertedRows = steps.map((step) => {
      if (step.kind === 'rowDelete') return [step.rowId];
      const current = working.parts.get(part.name)!,
        list = rows(current, step.tableId);
      const anchor = list.findIndex((row) => row.id === step.rowId);
      const span =
        step.where === 'above'
          ? list.slice(Math.max(0, anchor - step.count), anchor)
          : list.slice(anchor + 1, anchor + 1 + step.count);
      if (anchor < 0 || span.length !== step.count || span.some((row) => originalIds.has(row.id)))
        throw new Error('Native row allocation did not match the planned anchor.');
      return span.map((row) => row.id);
    });
    steps.forEach((step, index) => {
      if (step.seeds)
        working = materializeInsertedRow(
          working,
          part.name,
          step.tableId,
          insertedRows[index]![0]!,
          step.seeds,
          listContext
        );
    });
    const current = working.parts.get(part.name)!;
    const owners = new Map<string, number>();
    insertedRows.forEach((ids, index) => ids.forEach((id) => owners.set(id, index)));
    const addresses = steps.map(() => new Map<string, RevisionAddress>());
    for (const site of collectRevisionSites(current)) {
      const before = originalSites.get(site.node.id);
      if (before !== undefined) {
        if (before !== JSON.stringify(site.node))
          throw new Error('A row proposal changed an unrelated pending revision.');
        continue;
      }
      const row = ownerRow(current, site.node),
        index = row ? owners.get(row.id) : undefined,
        address = nativePropertyAddress(site.node);
      if (index === undefined || !address)
        throw new Error('A new native revision escaped its planned row.');
      addresses[index]!.set(propertyAddressKey(part.name, address), Object.freeze({ ...address }));
    }
    proposals = Object.freeze(
      steps.map((step, index) => {
        const group = [...addresses[index]!.values()];
        if (!group.length) throw new Error('A planned row did not receive a native revision.');
        return Object.freeze({
          identity: JSON.stringify([
            part.name,
            group.map((address) => propertyAddressKey(part.name, address)).sort(),
          ]),
          partName: part.name,
          kind: step.kind,
          tableId: step.tableId,
          rowIds: Object.freeze(insertedRows[index]!),
          addresses: Object.freeze(group),
        });
      })
    );
    return working;
  };
  const ops = steps.map((step) => step.op);
  const refused = preflightTrackedStructure(read.input, ops, complete, 'row', () => proposals);
  if (refused) return refused;
  proposals = [];
  const committed = commitTrackedStructure(read.input, ops, complete, 'row', () => proposals);
  if (!committed.success) return committed;
  return { success: true, revision: committed.revision, proposals };
}
