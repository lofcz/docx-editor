import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, findNode, paragraphTextOf } from '@docx-editor.dev/core/store';
import { manualColumnDocument, columnShape } from '../../test/support/manual-table-columns';
import { selectTableCells, savedTableFingerprint } from '../../test/support/manual-table-controls';
import {
  mergeManualTableCells,
  unmergeManualTableCell,
  createTableMergeToolbar,
} from './table-merges';

for (const direction of ['horizontal', 'vertical', 'rectangle'] as const)
  test(`${direction} merge and existing-grid split preserve original slots, all text and exact history`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: manualColumnDocument('none'),
    });
    const toolbar = createTableMergeToolbar(editor);
    try {
      const s = editor.surface!,
        source = columnShape(s.session.part()),
        rowEnd = direction === 'horizontal' ? 1 : 2,
        colEnd = direction === 'vertical' ? 0 : 1;
      const selected = selectTableCells(editor, [1, rowEnd], [0, colEnd]);
      const texts = source.rows
        .slice(1, rowEnd + 1)
        .flatMap((r) => r.cells.slice(0, colEnd + 1))
        .flatMap((c) => c.paragraphs);
      const before = canonicalOoxmlFingerprint(s.session.part());
      const merge = toolbar.element.querySelectorAll('button')[0]!;
      expect(merge.disabled).toBe(false);
      const event = new MouseEvent('mousedown', { button: 0, bubbles: true, cancelable: true });
      merge.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(true);
      merge.click();
      const merged = canonicalOoxmlFingerprint(s.session.part());
      for (const p of texts) expect(paragraphTextOf(s.session.part(), p.id)).toBe(p.text);
      expect(columnShape(s.session.part()).widths).toEqual([1200, 1800, 2400]);
      expect(await savedTableFingerprint(editor)).toEqual(merged);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(merged);
      const point = { paragraphId: texts[0]!.id, offset: 0 };
      s.setCellSelection(null);
      s.setSelection({ anchor: point, head: point });
      const split = toolbar.element.querySelectorAll('button')[1]!;
      expect(split.disabled).toBe(false);
      split.click();
      const after = columnShape(s.session.part());
      expect(after.widths).toEqual([1200, 1800, 2400]);
      const restored = after.rows[1]!.cells[0]!;
      expect(restored.span).toBe(1);
      expect(restored.width).toBe(1200);
      expect(restored.paragraphs.map((p) => p.text)).toEqual(texts.map((p) => p.text));
      for (const row of after.rows.slice(1, rowEnd + 1))
        for (const cell of row.cells.slice(0, colEnd + 1)) expect(cell.span).toBe(1);
      const splitFingerprint = canonicalOoxmlFingerprint(s.session.part());
      expect(await savedTableFingerprint(editor)).toEqual(splitFingerprint);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(merged);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(splitFingerprint);
      expect(findNode(s.session.part(), selected.tableId)?.kind).toBe('table');
    } finally {
      toolbar.dispose();
      editor.destroy();
    }
  });
test('invalid overlap refuses without native writes and displays the native reason', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: manualColumnDocument('horizontal'),
  });
  const toolbar = createTableMergeToolbar(editor);
  try {
    const s = editor.surface!,
      selection = selectTableCells(editor, [0, 1], [0, 1]);
    const shape = columnShape(s.session.part());
    const overlap = [shape.rows[0]!.cells[1]!.id, shape.rows[1]!.cells[1]!.id];
    expect(selection.cellIds.length).toBeGreaterThan(1);
    const before = canonicalOoxmlFingerprint(s.session.part()),
      revision = s.session.packageRevision();
    const refused = mergeManualTableCells(editor, overlap);
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.reason).toBe('cells-not-rectangle');
    expect(unmergeManualTableCell(editor).ok).toBe(false);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(s.session.packageRevision()).toBe(revision);
    expect(toolbar.element.querySelector('output')!.textContent!.length).toBeGreaterThan(0);
  } finally {
    toolbar.dispose();
    editor.destroy();
  }
});

test('a valid unmerge refuses in a viewing editor without changing the package', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: manualColumnDocument('horizontal'),
    mode: 'view',
  });
  try {
    const surface = editor.surface!,
      shape = columnShape(surface.session.part()),
      point = { paragraphId: shape.rows[0]!.cells[0]!.paragraphs[0]!.id, offset: 0 };
    surface.setSelection({ anchor: point, head: point });
    const before = canonicalOoxmlFingerprint(surface.session.part());
    const result = unmergeManualTableCell(editor);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.code).toBe('locked');
    expect(canonicalOoxmlFingerprint(surface.session.part())).toEqual(before);
  } finally {
    editor.destroy();
  }
});
