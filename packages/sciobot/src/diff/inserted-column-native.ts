import {
  findNode,
  readTwipsMeasure,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { InsertedTableColumnIntent } from './inserted-column-types';

const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
const wordChild = (node: OoxmlNode, name: string) =>
  children(node).find(
    (child) =>
      child.kind !== 'textValue' &&
      child.namespaceUri === WML_NAMESPACE_URI &&
      child.localName === name
  );
export function nativeGridColumns(part: OoxmlPart, tableId: string): readonly OoxmlNode[] {
  const table = findNode(part, tableId),
    grid = table && wordChild(table, 'tblGrid');
  return grid
    ? children(grid).filter(
        (node) =>
          node.kind !== 'textValue' &&
          node.namespaceUri === WML_NAMESPACE_URI &&
          node.localName === 'gridCol'
      )
    : [];
}

export function insertedColumnOperation(
  intent: InsertedTableColumnIntent
): Extract<TreeDocOp, { op: 'insertTableColumn' }> {
  return {
    op: 'insertTableColumn',
    tableId: intent.tableId,
    where: intent.where,
    ...(intent.anchor.kind === 'grid-column'
      ? { gridColumnId: intent.anchor.id }
      : { referenceCellId: intent.anchor.id }),
  };
}

/** Resolve a just-created native grid column. No allocated IDs belong in the serialized intent. */
export function insertedColumnIndex(
  before: OoxmlPart,
  after: OoxmlPart,
  intent: InsertedTableColumnIntent
): number | null {
  const previous = nativeGridColumns(before, intent.tableId),
    next = nativeGridColumns(after, intent.tableId);
  if (intent.anchor.kind === 'grid-column') {
    const anchor = next.findIndex((column) => column.id === intent.anchor.id);
    if (anchor < 0 || next.length !== previous.length + 1) return null;
    const index = anchor + (intent.where === 'left' ? -1 : 1),
      column = next[index];
    return column && !previous.some((entry) => entry.id === column.id) ? index : null;
  }
  const table = findNode(before, intent.tableId),
    firstRow = table && children(table).find((node) => node.kind === 'tableRow');
  const cells = firstRow ? children(firstRow).filter((node) => node.kind === 'tableCell') : [];
  const anchor = cells.findIndex((cell) => cell.id === intent.anchor.id);
  if (previous.length || anchor < 0 || next.length !== cells.length + 1) return null;
  return anchor + (intent.where === 'left' ? 0 : 1);
}

/** Use the native fixed-grid sizing command. Preserve all other authored grid widths. */
export function insertedColumnWidthOperation(
  before: OoxmlPart,
  after: OoxmlPart,
  intent: InsertedTableColumnIntent
):
  | {
      readonly ok: true;
      readonly operation: Extract<TreeDocOp, { op: 'setTableProperties' }> | null;
    }
  | { readonly ok: false; readonly reason: string } {
  if (intent.properties.widthTwips === undefined) return { ok: true, operation: null };
  const index = insertedColumnIndex(before, after, intent);
  if (index === null)
    return { ok: false, reason: 'New native grid column is missing or ambiguous.' };
  const widths: number[] = [];
  for (const node of nativeGridColumns(after, intent.tableId)) {
    const raw =
      node.kind === 'textValue'
        ? undefined
        : node.attributes.find(
            (attribute) =>
              attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'w'
          )?.value;
    const width = widths.length === index ? intent.properties.widthTwips : readTwipsMeasure(raw);
    if (width === null)
      return { ok: false, reason: 'Explicit grid sizing requires readable existing widths.' };
    widths.push(width);
  }
  return {
    ok: true,
    operation: { op: 'setTableProperties', tableId: intent.tableId, columnWidthsTwips: widths },
  };
}
