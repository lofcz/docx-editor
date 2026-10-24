import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ExecResult } from '@docx-editor.dev/core/contracts/editor';
import { parentNodeOf, WML_NAMESPACE_URI } from '@docx-editor.dev/core/store';
import {
  paragraphsInCells,
  tableContextAt,
  type BlockFragmentRecord,
} from '@docx-editor.dev/core/layout';

/** Native table insertion allocates real Word cells and adopts the first cell's caret. */
export function insertManualTable(
  editor: DocxEditorInstance,
  rows: number,
  columns: number
): ExecResult {
  editor.surface?.flushPendingInput();
  return editor.exec({ type: 'insertTable', rows, cols: columns });
}

/** Read logical cells once, skipping repeated page headers and continued cell fragments. */
export function nativeTableCellIds(editor: DocxEditorInstance, tableId: string): readonly string[] {
  const result: string[] = [],
    seen = new Set<string>();
  const visit = (blocks: readonly BlockFragmentRecord[]) => {
    for (const block of blocks) {
      if (block.kind !== 'table') continue;
      for (const row of block.rows) {
        if (row.isHeaderRepeat) continue;
        for (const cell of row.cells) {
          if (block.tableId === tableId && !seen.has(cell.id)) {
            seen.add(cell.id);
            result.push(cell.id);
          }
          visit(cell.blocks);
        }
      }
    }
  };
  for (const page of editor.surface?.layout().pages ?? []) visit(page.fragments);
  return result;
}

/** Resolve the active native cell using the public layout projection. */
export function nativeTableAnchor(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) return null;
  const layout = surface.layout(),
    paragraph = surface.state().selection.head.paragraphId;
  const context = tableContextAt(layout, paragraph);
  if (!context) return null;
  const rectangle = editor.getTableCellSelection();
  if (!rectangle) {
    const part = surface.session.part();
    let parent = parentNodeOf(part, paragraph);
    for (let depth = 0; parent && depth < 256; depth++) {
      if (parent.namespaceUri === WML_NAMESPACE_URI && parent.localName === 'tc')
        return { tableId: context.tableId, cellId: parent.id };
      parent = parentNodeOf(part, parent.id);
    }
  }
  const cellId =
    rectangle?.tableId === context.tableId
      ? rectangle.cellIds[0]
      : nativeTableCellIds(editor, context.tableId).find((cell) =>
          paragraphsInCells(layout, [cell]).includes(paragraph)
        );
  return cellId ? { tableId: context.tableId, cellId } : null;
}

/** App table navigation uses native model selections; it never reads browser text positions. */
export function navigateTableCell(
  editor: DocxEditorInstance,
  direction: 'next' | 'previous'
): ExecResult {
  if (direction !== 'next' && direction !== 'previous')
    return { ok: false, code: 'invalidArgs', reason: 'Choose a table navigation direction.' };
  const surface = editor.surface;
  if (!surface) return { ok: false, code: 'locked', reason: 'The native editor is not attached.' };
  surface.flushPendingInput();
  const anchor = nativeTableAnchor(editor);
  if (!anchor) return { ok: false, code: 'unsupported', reason: 'Select a native table cell.' };
  let cells = nativeTableCellIds(editor, anchor.tableId);
  const index = cells.indexOf(anchor.cellId);
  if (index < 0)
    return { ok: false, code: 'notFound', reason: 'The selected table cell is not placed.' };
  const next = index + (direction === 'next' ? 1 : -1);
  let changed = false;
  if (next < 0) return { ok: true, changed: false };
  if (next === cells.length) {
    const inserted = editor.exec({ type: 'insertRow', where: 'below' });
    if (!inserted.ok) return inserted;
    changed = inserted.changed;
    cells = nativeTableCellIds(editor, anchor.tableId);
  }
  const cell = cells[next];
  const paragraphId = cell && paragraphsInCells(surface.layout(), [cell])[0];
  if (!paragraphId)
    return { ok: false, code: 'notFound', reason: 'The destination cell has no paragraph.' };
  const point = { paragraphId, offset: 0 };
  surface.setSelection({ anchor: point, head: point });
  surface.focus();
  return { ok: true, changed };
}

/** Install the application's Tab behavior ahead of the engine's ordinary paragraph keymap. */
export function installTableNavigation(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const host = surface.hostElement();
  const keydown = (event: KeyboardEvent) => {
    const target = event.target as Element | null;
    if (
      event.key !== 'Tab' ||
      event.isComposing ||
      event.ctrlKey ||
      event.metaKey ||
      event.altKey ||
      !target?.closest?.('[contenteditable="true"]') ||
      target.closest('[data-docx-block-ui]') ||
      surface.contentControls.formFill()
    )
      return;
    if (!nativeTableAnchor(editor)) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    navigateTableCell(editor, event.shiftKey ? 'previous' : 'next');
  };
  host.addEventListener('keydown', keydown, true);
  return () => host.removeEventListener('keydown', keydown, true);
}

/** The production picker offers an eight-by-eight grid. A pointer pick retains native selection. */
export function createTableInsertionGrid(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'grid');
  element.setAttribute('aria-label', 'Insert table');
  const current = element.ownerDocument.createElement('output');
  current.setAttribute('aria-label', 'Table size');
  element.append(current);
  const buttons: { button: HTMLButtonElement; rows: number; columns: number }[] = [];
  for (let rows = 1; rows <= 8; rows++) {
    const row = element.ownerDocument.createElement('div');
    row.setAttribute('role', 'row');
    for (let columns = 1; columns <= 8; columns++) {
      const cell = element.ownerDocument.createElement('div');
      cell.setAttribute('role', 'gridcell');
      const button = element.ownerDocument.createElement('button');
      button.type = 'button';
      button.textContent = `${rows}×${columns}`;
      button.setAttribute('aria-label', `Insert ${rows} by ${columns} table`);
      button.addEventListener('mousedown', (event) => {
        if (event.button === 0) event.preventDefault();
      });
      button.addEventListener('pointerenter', () => {
        current.textContent = `${rows} × ${columns}`;
        for (const entry of buttons)
          entry.button.setAttribute(
            'aria-pressed',
            String(entry.rows <= rows && entry.columns <= columns)
          );
      });
      button.addEventListener('click', () => {
        if (insertManualTable(editor, rows, columns).ok) surface.focus();
      });
      cell.append(button);
      row.append(cell);
      buttons.push({ button, rows, columns });
    }
    element.append(row);
  }
  const update = () => {
    for (const { button, rows, columns } of buttons) {
      const allowed = editor.can({ type: 'insertTable', rows, cols: columns });
      button.disabled = !allowed.ok;
      button.title = allowed.ok ? '' : allowed.reason;
    }
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
