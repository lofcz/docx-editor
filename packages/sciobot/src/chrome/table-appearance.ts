import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { TreeDocOp, OoxmlNode } from '@docx-editor.dev/core/store';
import { WML_NAMESPACE_URI } from '@docx-editor.dev/core/store';
import { nativeGridColumns } from '../diff/inserted-column-native';
import {
  applyManualTableOps,
  canManualTableOps,
  manualTableContext,
  manualTableAddresses,
  tableRefusal,
} from './table-native';

export type TableAppearanceChange =
  | { kind: 'fill'; color: string | null }
  | { kind: 'column-width'; twips: number }
  | { kind: 'row-height'; twips: number }
  | { kind: 'headers'; count: number }
  | { kind: 'borders'; color: string; size: number }
  | { kind: 'clear-borders' };
const attr = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)
        ?.value;
function operation(editor: DocxEditorInstance, change: TableAppearanceChange): TreeDocOp | null {
  const context = manualTableContext(editor);
  if (!context) return null;
  const { part, table, address } = context;
  const cellIds = editor.getTableCellSelection()?.cellIds ?? [address.cellId];
  switch (change.kind) {
    case 'fill':
      return {
        op: 'setTableCellFill',
        tableId: table.id,
        cellIds,
        color:
          change.color === null
            ? null
            : { kind: 'hex', value: change.color.replace(/^#/, '').toUpperCase() },
      };
    case 'column-width': {
      const columns = nativeGridColumns(part, table.id),
        widths = columns.map((c) => Number(attr(c, 'w')));
      const index =
        manualTableAddresses(editor, table.id).get(address.cellId)?.logicalGridColumn ??
        address.gridColumn;
      if (!columns[index]) return null;
      widths[index] = change.twips;
      return { op: 'setTableProperties', tableId: table.id, columnWidthsTwips: widths };
    }
    case 'row-height':
      return {
        op: 'setTableRowHeight',
        tableId: table.id,
        rowId: address.rowId,
        heightTwips: change.twips,
      };
    case 'headers':
      return { op: 'setTableProperties', tableId: table.id, headerRowCount: change.count };
    case 'borders':
      return {
        op: 'setTableCellBorders',
        tableId: table.id,
        cellIds,
        scope: 'all',
        spec: {
          style: 'single',
          size: change.size,
          color: { kind: 'hex', value: change.color.replace(/^#/, '').toUpperCase() },
        },
      };
    case 'clear-borders':
      return {
        op: 'setTableCellBorders',
        tableId: table.id,
        cellIds,
        scope: 'none',
        target: 'all',
      };
  }
}
export function canChangeManualTableAppearance(
  editor: DocxEditorInstance,
  change: TableAppearanceChange
) {
  const op = operation(editor, change);
  return op ? canManualTableOps(editor, [op]) : tableRefusal('Select a native table cell.');
}
export function changeManualTableAppearance(
  editor: DocxEditorInstance,
  change: TableAppearanceChange
) {
  editor.surface?.flushPendingInput();
  const op = operation(editor, change);
  return op ? applyManualTableOps(editor, [op]) : tableRefusal('Select a native table cell.');
}
export function createTableAppearanceToolbar(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Table appearance');
  const input = (label: string, type: string, value: string) => {
    const e = element.ownerDocument.createElement('input');
    e.type = type;
    e.value = value;
    e.setAttribute('aria-label', label);
    element.append(e);
    return e;
  };
  const color = input('Cell background', 'color', '#ffee00'),
    width = input('Column width (pt)', 'number', '90'),
    height = input('Row height (pt)', 'number', '30'),
    headers = input('Header rows', 'number', '1');
  const actions: [string, () => TableAppearanceChange][] = [
    ['Apply cell background', () => ({ kind: 'fill', color: color.value })],
    ['Clear cell background', () => ({ kind: 'fill', color: null })],
    ['Apply column width', () => ({ kind: 'column-width', twips: Number(width.value) * 20 })],
    ['Apply row height', () => ({ kind: 'row-height', twips: Number(height.value) * 20 })],
    ['Apply header rows', () => ({ kind: 'headers', count: Number(headers.value) })],
    ['All borders', () => ({ kind: 'borders', color: '000000', size: 4 })],
    ['Clear borders', () => ({ kind: 'clear-borders' })],
  ];
  const status = element.ownerDocument.createElement('output');
  status.setAttribute('role', 'status');
  const buttons = actions.map(([label, change]) => {
    const button = element.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.addEventListener('mousedown', (e) => {
      if (e.button === 0) e.preventDefault();
    });
    button.addEventListener('click', () => {
      const result = changeManualTableAppearance(editor, change());
      if (result.ok) {
        surface.focus();
        update();
      } else status.textContent = result.reason;
    });
    element.append(button);
    return { button, change };
  });
  const update = () => {
    for (const { button, change } of buttons) {
      const state = canChangeManualTableAppearance(editor, change());
      button.disabled = !state.ok;
      button.title = state.ok ? '' : state.reason;
    }
  };
  for (const e of [color, width, height, headers]) e.addEventListener('input', update);
  element.append(status);
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
