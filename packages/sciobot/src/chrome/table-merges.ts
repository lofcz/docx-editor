import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { TreeDocOp } from '@docx-editor.dev/core/store';
import { cellSelectionBetween } from '@docx-editor.dev/core/layout';
import {
  applyManualTableOps,
  canManualTableOps,
  manualTableContext,
  tableRefusal,
} from './table-native';

function splitOperation(editor: DocxEditorInstance): TreeDocOp | null {
  const context = manualTableContext(editor);
  if (!context) return null;
  const { surface, table, address } = context;
  const extent = cellSelectionBetween(surface.layout(), address, address);
  if (!extent) return null;
  const selected = editor.getTableCellSelection();
  if (selected?.cellIds.some((id) => !extent.cellIds.includes(id))) return null;
  const rows = extent.rows.to - extent.rows.from + 1,
    columns = extent.columns.to - extent.columns.from + 1;
  if (rows * columns < 2) return null;
  return {
    op: 'splitTableCell',
    tableId: table.id,
    cellId: address.cellId,
    rows,
    columns,
    columnDistribution: 'existing-grid',
  };
}
export function canUnmergeManualTableCell(editor: DocxEditorInstance) {
  const op = splitOperation(editor);
  return op
    ? canManualTableOps(editor, [op], true)
    : tableRefusal('Select one merged cell to split.');
}
/** Restore the original Word slots, including unequal columns; retain content in cell one. */
export function unmergeManualTableCell(editor: DocxEditorInstance) {
  editor.surface?.flushPendingInput();
  const op = splitOperation(editor);
  return op
    ? applyManualTableOps(editor, [op], true)
    : tableRefusal('Select one merged cell to split.');
}
export function mergeManualTableCells(editor: DocxEditorInstance, cellIds?: readonly string[]) {
  editor.surface?.flushPendingInput();
  if (cellIds === undefined) return editor.exec({ type: 'mergeCells' });
  const context = manualTableContext(editor);
  return context
    ? applyManualTableOps(
        editor,
        [{ op: 'mergeTableCells', tableId: context.table.id, cellIds }],
        true
      )
    : tableRefusal('Select a native table cell.');
}
export function createTableMergeToolbar(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Table cells');
  const status = element.ownerDocument.createElement('output');
  status.setAttribute('role', 'status');
  const actions = [
    {
      label: 'Merge cells',
      can: () => editor.can({ type: 'mergeCells' }),
      run: () => mergeManualTableCells(editor),
    },
    {
      label: 'Split cell',
      can: () => canUnmergeManualTableCell(editor),
      run: () => unmergeManualTableCell(editor),
    },
  ];
  const buttons = actions.map((action) => {
    const button = element.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent = action.label;
    button.addEventListener('mousedown', (e) => {
      if (e.button === 0) e.preventDefault();
    });
    button.addEventListener('click', () => {
      const result = action.run();
      if (result.ok) {
        surface.focus();
        update();
      } else status.textContent = result.reason;
    });
    element.append(button);
    return { button, action };
  });
  element.append(status);
  const update = () => {
    const reasons: string[] = [];
    for (const { button, action } of buttons) {
      const state = action.can();
      button.disabled = !state.ok;
      button.title = state.ok ? '' : state.reason;
      if (!state.ok) reasons.push(`${action.label}: ${state.reason}`);
    }
    status.textContent = reasons.join(' ');
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
