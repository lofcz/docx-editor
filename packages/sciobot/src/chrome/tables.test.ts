import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { paragraphsInCells } from '@docx-editor.dev/core/layout';
import { docxFixture } from '../../test/support/document';
import reference from '../__tests__/fixtures/native-commands.json';
import {
  createTableInsertionGrid,
  insertManualTable,
  installTableNavigation,
  nativeTableCellIds,
  nativeTableAnchor,
  navigateTableCell,
} from './tables';

test('2x3 insertion keeps source text, adopts cell one, supports native cell navigation and exact one-step undo', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
    ),
  });
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!,
      at = { paragraphId: id, offset: 5 };
    s.setSelection({ anchor: at, head: at });
    const before = canonicalOoxmlFingerprint(s.session.part());
    expect(reference.cases.find((c) => c.id === 'table.insert')).toBeDefined();
    expect(insertManualTable(editor, 2, 3).ok).toBe(true);
    const anchor = nativeTableAnchor(editor)!;
    expect(anchor).not.toBeNull();
    const cells = nativeTableCellIds(editor, anchor.tableId);
    expect(cells).toHaveLength(6);
    expect(anchor.cellId).toBe(cells[0]!);
    expect(s.state().selection.head.offset).toBe(0);
    expect(paragraphTextOf(s.session.part(), id)).toBe('Alpha beta');
    expect(navigateTableCell(editor, 'next')).toEqual({ ok: true, changed: false });
    expect(nativeTableAnchor(editor)?.cellId).toBe(cells[1]!);
    expect(navigateTableCell(editor, 'previous').ok).toBe(true);
    expect(nativeTableAnchor(editor)?.cellId).toBe(cells[0]!);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    expect(
      canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
    ).toEqual(canonicalOoxmlFingerprint(s.session.part()));
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(nativeTableCellIds(editor, anchor.tableId)).toHaveLength(6);
  } finally {
    editor.destroy();
  }
});

test('grid pointer keeps the insertion caret, Tab/Shift+Tab navigate, and last-cell Tab appends one native row', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docxFixture('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>'),
  });
  const grid = createTableInsertionGrid(editor),
    stop = installTableNavigation(editor);
  try {
    expect(grid.element.querySelectorAll('[role="gridcell"]')).toHaveLength(64);
    const button = grid.element.querySelector('[aria-label="Insert 2 by 3 table"]')!;
    const event = new MouseEvent('mousedown', { button: 0, bubbles: true, cancelable: true });
    button.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
    (button as HTMLButtonElement).click();
    const s = editor.surface!,
      anchor = nativeTableAnchor(editor)!;
    const cells = nativeTableCellIds(editor, anchor.tableId);
    const editable = s.hostElement().querySelector('[contenteditable="true"]')!;
    const tab = (shiftKey = false) => {
      const event = new KeyboardEvent('keydown', {
        key: 'Tab',
        code: 'Tab',
        shiftKey,
        bubbles: true,
        cancelable: true,
      });
      editable.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(true);
    };
    tab();
    expect(nativeTableAnchor(editor)?.cellId).toBe(cells[1]!);
    tab(true);
    expect(nativeTableAnchor(editor)?.cellId).toBe(cells[0]!);
    const last = paragraphsInCells(s.layout(), [cells[5]!])[0]!,
      point = { paragraphId: last, offset: 0 };
    s.setSelection({ anchor: point, head: point });
    const before = canonicalOoxmlFingerprint(s.session.part());
    tab();
    expect(nativeTableCellIds(editor, anchor.tableId)).toHaveLength(9);
    expect(nativeTableAnchor(editor)?.cellId).toBe(nativeTableCellIds(editor, anchor.tableId)[6]!);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    stop();
  } finally {
    stop();
    grid.dispose();
    editor.destroy();
    container.remove();
  }
});

test('invalid dimensions refuse without changing native selection, package, or history', () => {
  const editor = createDocxEditor({ container: document.createElement('div'), document: 'blank' });
  try {
    const s = editor.surface!,
      before = canonicalOoxmlFingerprint(s.session.part()),
      selection = s.state().selection;
    for (const [rows, cols] of [
      [0, 3],
      [2, 0],
      [1.5, 3],
      [1000000, 1000000],
    ]) {
      expect(insertManualTable(editor, rows!, cols!).ok).toBe(false);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(s.state().selection).toEqual(selection);
      expect(s.state().canUndo).toBe(false);
    }
  } finally {
    editor.destroy();
  }
});
