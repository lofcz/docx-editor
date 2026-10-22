import { findNode, type TransactionRevisionIds } from '../package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import type { TreeDocOp } from './tree-op-types.ts';

/** Native-generated structural sites owned by one part in one open store transaction. */
const states = new WeakMap<TransactionRevisionIds['wrote'], Map<string, string>>();
const hidden = new WeakMap<TransactionRevisionIds['wrote'], ReadonlySet<string>>();
export const TABLE_RECORDS = new Set([
  'ins',
  'del',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
  'trPrChange',
  'tblPrChange',
  'tblPrExChange',
  'tblGridChange',
]);
export const TABLE_CONTAINERS = new Set(['tblPr', 'tblGrid', 'trPr', 'tblPrEx', 'tcPr']);
export const word = (node: OoxmlNode, name: string): node is OoxmlElement =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;
export const kids = (node: OoxmlNode | undefined): readonly OoxmlNode[] =>
  node && node.kind !== 'textValue' ? node.children : [];
export const child = (node: OoxmlNode | undefined, name: string): OoxmlElement | undefined =>
  kids(node).find((item): item is OoxmlElement => word(item, name));
export function isTrackedTableOp(op: TreeDocOp): boolean {
  return (
    'revision' in op &&
    !!op.revision &&
    ('tableId' in op || (op.op === 'authorTable' && op.action.kind === 'existing'))
  );
}
export function trackedTableId(op: TreeDocOp): string | null {
  if (!isTrackedTableOp(op)) return null;
  return 'tableId' in op
    ? op.tableId
    : op.op === 'authorTable' && op.action.kind === 'existing'
      ? op.action.tableId
      : null;
}
/** Direct containers only: history copies and nested tables are separate owners. */
export function tableContainers(table: OoxmlNode): readonly OoxmlElement[] {
  const result: OoxmlElement[] = [];
  const append = (node: OoxmlNode) => {
    for (const item of kids(node))
      if (
        item.kind !== 'textValue' &&
        TABLE_CONTAINERS.has(item.localName) &&
        item.namespaceUri === WML_NAMESPACE_URI
      )
        result.push(item);
  };
  append(table);
  for (const row of kids(table))
    if (row.kind === 'tableRow') {
      append(row);
      for (const cell of kids(row)) if (cell.kind === 'tableCell') append(cell);
    }
  return result;
}
export const tableRecords = (table: OoxmlNode) =>
  tableContainers(table).flatMap((container) =>
    kids(container).filter(
      (node): node is OoxmlElement =>
        node.kind !== 'textValue' &&
        node.namespaceUri === WML_NAMESPACE_URI &&
        TABLE_RECORDS.has(node.localName)
    )
  );
export function ownedTableSites(ids: TransactionRevisionIds): ReadonlyMap<string, string> {
  return states.get(ids.wrote) ?? new Map();
}
export function rememberTableSites(
  before: OoxmlPart,
  after: OoxmlPart,
  op: TreeDocOp,
  ids: TransactionRevisionIds | undefined
): void {
  const tableId = trackedTableId(op);
  if (!ids || !tableId) return;
  const old = findNode(before, tableId),
    next = findNode(after, tableId);
  if (!old || !next) return;
  const original = new Set(tableRecords(old).map((node) => node.id));
  let owned = states.get(ids.wrote);
  if (!owned) {
    owned = new Map();
    states.set(ids.wrote, owned);
  }
  for (const node of tableRecords(next))
    if (!original.has(node.id) && !owned.has(node.id)) owned.set(node.id, op.op);
}
export function withHiddenTableParagraphs<T>(
  ids: TransactionRevisionIds,
  paragraphs: ReadonlySet<string>,
  run: () => T
): T {
  const previous = hidden.get(ids.wrote);
  hidden.set(ids.wrote, paragraphs);
  try {
    return run();
  } finally {
    if (previous) hidden.set(ids.wrote, previous);
    else hidden.delete(ids.wrote);
  }
}
export const isHiddenTableParagraph = (ids: TransactionRevisionIds | undefined, id: string) =>
  !!ids && !!hidden.get(ids.wrote)?.has(id);
