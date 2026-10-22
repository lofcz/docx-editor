import { cellPropertiesAt } from '../diff/new-cell-properties';
import type { DirectListContext } from '../ai/direct-lists';
import { cellListBindings } from '../diff/inserted-cell-lists';
import { resolveTableGeometry, geometrySplitCells } from '../diff/table-geometry';
import {
  collectRevisionSites,
  findNode,
  parentNodeOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
  type RevisionAddress,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit } from '../diff/plan';
import { insertedColumnIndex, insertedColumnOperation } from '../diff/inserted-column-native';
import { nativeVisualSize, splitResultCells } from '../diff/split-cell-native';
import type { ReviewGroupReceipt } from './groups';
import {
  nativePropertyAddress,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import { stageNewCells } from './tracked-stage-cells';
import { currentStagePart } from './tracked-stage-text';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

type TableEdit = Extract<
  CanonicalEdit,
  { kind: 'table' | 'insert-row' | 'insert-column' | 'split-cell' | 'table-geometry' }
>;
const PROPERTY_OPS = new Set([
  'setTableProperties',
  'setTableColumnWidths',
  'setTableRightEdgeWidth',
  'setTableRowHeight',
  'setTableCellBorders',
  'setTableCellFill',
  'setTableCellVerticalAlignment',
]);
const STRUCTURE_OPS = new Set([
  'insertTableRow',
  'deleteTableRow',
  'insertTableColumn',
  'deleteTableColumn',
  'mergeTableCells',
  'splitTableCell',
]);
const children = (node: OoxmlNode | null | undefined): readonly OoxmlNode[] =>
  node && node.kind !== 'textValue' ? node.children : [];
const rows = (part: OoxmlPart, tableId: string) =>
  children(findNode(part, tableId)).filter((node) => node.kind === 'tableRow');
const cells = (part: OoxmlPart, rowId: string) =>
  children(findNode(part, rowId)).filter((node) => node.kind === 'tableCell');
export const trackedTableId = (edit: TableEdit): string => {
  if (edit.kind !== 'table') return edit.intent.tableId;
  const op = edit.operation;
  if ('tableId' in op) return op.tableId;
  if (op.op === 'authorTable' && op.action.kind === 'existing') return op.action.tableId;
  return refuseWrite(
    'Creating a complete tracked table needs its fragment stage.',
    'unsupported-operation'
  );
};
export const tableStageHasStructure = (edit: TableEdit): boolean =>
  edit.kind !== 'table' || !PROPERTY_OPS.has(edit.operation.op);
function tableOwner(part: OoxmlPart, node: OoxmlNode): string | null {
  let current: OoxmlNode | null = node;
  for (let depth = 0; current && depth < 64; depth++) {
    if (current.kind === 'table') return current.id;
    current = parentNodeOf(part, current.id);
  }
  return null;
}
function address(node: OoxmlNode): RevisionAddress | null {
  if (node.kind === 'textValue') return null;
  if (node.localName !== 'tblGridChange') return nativePropertyAddress(node);
  const id = node.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'id'
  )?.value;
  return id === undefined ? null : { id, author: '' };
}

/** Lower one ordered table intent, then collect only revisions that this step actually creates. */
export function stageTrackedTable(
  writer: TrackedStageWriter,
  edit: TableEdit,
  revision: TrackedStageAttribution,
  trackFormatting: boolean,
  listContext?: DirectListContext
): readonly ReviewGroupReceipt[] {
  const before = currentStagePart(writer, edit.partName),
    tableId = trackedTableId(edit);
  if (findNode(before, tableId)?.kind !== 'table')
    refuseWrite('A planned table no longer exists.', 'stale-revision');
  const protectedContent = propertyControlRefusal(before, tableId);
  if (protectedContent) refuseWrite(protectedContent, 'unsupported-operation');
  const original = new Map(
    collectRevisionSites(before).map((site) => [site.node.id, JSON.stringify(site.node)])
  );
  const apply = (op: TreeDocOp) => writer.apply(edit.partName, op);
  if (edit.kind === 'table') {
    const op = edit.operation;
    if (PROPERTY_OPS.has(op.op) && !trackFormatting)
      refuseWrite(
        'Formatting tracking is disabled for this document or editor.',
        'unsupported-operation'
      );
    if (
      !PROPERTY_OPS.has(op.op) &&
      !STRUCTURE_OPS.has(op.op) &&
      !(
        op.op === 'authorTable' &&
        op.action.kind === 'existing' &&
        op.action.mutation.kind === 'insertRows'
      )
    )
      refuseWrite('This table command needs its tracked fragment stage.', 'unsupported-operation');
    apply({ ...op, revision } as TreeDocOp);
  } else if (edit.kind === 'table-geometry') {
    const resolved = resolveTableGeometry(before, edit.intent);
    if (!resolved.success) refuseWrite(resolved.error, resolved.code);
    apply({ ...resolved.operation, revision } as TreeDocOp);
    if (edit.intent.kind === 'split') {
      const fresh = geometrySplitCells(
        currentStagePart(writer, edit.partName),
        edit.intent,
        resolved.cellId
      );
      if (!fresh || fresh.length !== edit.intent.cells.length)
        refuseWrite('Native geometry split allocation differs from its plan.', 'apply-failed');
      const streams = edit.intent.cells;
      stageNewCells(
        writer,
        edit.partName,
        fresh.map((cellId, index) => ({
          cellId,
          blocks: streams[index]!,
          properties:
            edit.intent.kind === 'split'
              ? cellPropertiesAt(edit.intent.cellProperties, index)
              : undefined,
        }))
      );
    }
  } else if (edit.kind === 'insert-row') {
    const intent = edit.intent;
    if (intent.payloads.length)
      refuseWrite(
        'Application blocks in new rows need their tracked fragment stage.',
        'unsupported-operation'
      );
    const old = new Set(rows(before, tableId).map((row) => row.id));
    apply({ op: 'insertTableRow', tableId, rowId: intent.rowId, where: intent.where, revision });
    const part = currentStagePart(writer, edit.partName),
      list = rows(part, tableId),
      anchor = list.findIndex((row) => row.id === intent.rowId),
      index = anchor + (intent.where === 'above' ? -1 : 1),
      row = list[index];
    if (anchor < 0 || !row || old.has(row.id))
      refuseWrite('Native row allocation differs from its planned anchor.', 'apply-failed');
    const fresh = cells(part, row.id);
    if (fresh.length !== intent.cells.length)
      refuseWrite('New row cells differ from the planned physical cells.', 'apply-failed');
    if (intent.properties.heightTwips !== undefined)
      apply({
        op: 'setTableRowHeight',
        tableId,
        rowId: row.id,
        heightTwips: intent.properties.heightTwips,
      });
    stageNewCells(
      writer,
      edit.partName,
      fresh.map((cell, index) => ({
        cellId: cell.id,
        blocks: intent.cells[index]!,
        lists: cellListBindings(intent.lists, index),
        properties: cellPropertiesAt(intent.cellProperties, index),
      })),
      listContext
    );
  } else if (edit.kind === 'insert-column') {
    const intent = edit.intent;
    if (intent.payloads.length)
      refuseWrite(
        'Application blocks in new columns need their tracked fragment stage.',
        'unsupported-operation'
      );
    apply({
      ...insertedColumnOperation(intent),
      revision,
      ...(intent.properties.widthTwips === undefined
        ? {}
        : { widthTwips: intent.properties.widthTwips }),
    });
    const part = currentStagePart(writer, edit.partName);
    if (insertedColumnIndex(before, part, intent) === null)
      refuseWrite('Native column allocation is ambiguous.', 'apply-failed');
    const fresh = intent.cells.flatMap((entry, index) => {
      const old = new Set(cells(before, entry.rowId).map((cell) => cell.id)),
        created = cells(part, entry.rowId).filter((cell) => !old.has(cell.id));
      if (created.length > 1 || (!created.length && entry.blocks.length))
        refuseWrite('A new column cell is ambiguous.', 'apply-failed');
      return created.length
        ? [
            {
              cellId: created[0]!.id,
              blocks: entry.blocks,
              lists: cellListBindings(intent.lists, index),
              properties: cellPropertiesAt(intent.cellProperties, index),
            },
          ]
        : [];
    });
    stageNewCells(writer, edit.partName, fresh, listContext);
  } else {
    const intent = edit.intent;
    if (intent.payloads.length)
      refuseWrite(
        'Application blocks in split cells need their tracked fragment stage.',
        'unsupported-operation'
      );
    const size = nativeVisualSize(before, tableId, intent.cellId);
    if (!size) refuseWrite('The planned merged cell no longer exists.', 'stale-revision');
    apply({
      op: 'splitTableCell',
      tableId,
      cellId: intent.cellId,
      rows: intent.rows,
      columns: intent.columns,
      columnDistribution: 'existing-grid',
      revision,
    });
    const fresh = splitResultCells(
      currentStagePart(writer, edit.partName),
      tableId,
      intent.cellId,
      size
    );
    if (!fresh || fresh.length !== intent.cells.length)
      refuseWrite('The split cells differ from the planned region.', 'apply-failed');
    stageNewCells(
      writer,
      edit.partName,
      fresh.map((cellId, index) => ({
        cellId,
        blocks: intent.cells[index]!,
        properties: cellPropertiesAt(intent.cellProperties, index),
      }))
    );
  }
  const current = currentStagePart(writer, edit.partName),
    found = new Map<string, RevisionAddress>();
  for (const site of collectRevisionSites(current)) {
    const prior = original.get(site.node.id);
    if (prior !== undefined) {
      if (prior !== JSON.stringify(site.node))
        refuseWrite('A table stage changed an existing pending revision.', 'ambiguous-target');
      continue;
    }
    const receipt = address(site.node);
    if (tableOwner(current, site.node) !== tableId || !receipt || site.refused)
      refuseWrite('A new revision escaped its planned native table.', 'apply-failed');
    found.set(propertyAddressKey(edit.partName, receipt), receipt);
  }
  for (const id of original.keys())
    if (!findNode(current, id))
      refuseWrite('A table stage removed an existing pending revision.', 'ambiguous-target');
  return found.size ? [{ partName: edit.partName, addresses: [...found.values()] }] : [];
}

/** A structural table decision owns its fresh cell text and formatting dependencies too. */
export function coupleTableStageReceipts(
  writer: TrackedStageWriter,
  receipts: readonly ReviewGroupReceipt[],
  tables: ReadonlySet<string>,
  initial: ReadonlySet<string>
): readonly ReviewGroupReceipt[] {
  const grouped = new Map<string, { partName: string; addresses: Map<string, RevisionAddress> }>();
  const claimed = new Set<string>();
  for (const [partName, part] of writer.currentPackage().parts) {
    for (const site of collectRevisionSites(part)) {
      if (initial.has(site.node.id)) continue;
      const owner = tableOwner(part, site.node),
        key = JSON.stringify([partName, owner]);
      if (!owner || !tables.has(key)) continue;
      const found = address(site.node);
      if (!found || site.refused)
        refuseWrite('A table dependency has no native review address.', 'apply-failed');
      const entry = grouped.get(key) ?? { partName, addresses: new Map<string, RevisionAddress>() };
      const identity = propertyAddressKey(partName, found);
      entry.addresses.set(identity, found);
      claimed.add(identity);
      grouped.set(key, entry);
    }
  }
  return [
    ...receipts.flatMap((receipt) => {
      const remaining = receipt.addresses.filter(
        (item) => !claimed.has(propertyAddressKey(receipt.partName, item))
      );
      return remaining.length ? [{ ...receipt, addresses: remaining }] : [];
    }),
    ...Array.from(grouped.values(), (entry) => ({
      partName: entry.partName,
      addresses: [...entry.addresses.values()],
    })),
  ];
}
