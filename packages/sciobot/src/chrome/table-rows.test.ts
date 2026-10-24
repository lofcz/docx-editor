import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, findNode, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { paragraphsInCells } from '@docx-editor.dev/core/layout';
import { manualTableDocument } from '../../test/support/manual-table';
import reference from '../__tests__/fixtures/native-commands.json';
import { nativeTableAnchor, nativeTableCellIds } from './tables';
import { changeManualTableRow, createTableRowToolbar } from './table-rows';

for (const action of ['above', 'below', 'delete'] as const)
  test(`row ${action} preserves neighboring rows, header semantics, content, exact undo and reopen`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: manualTableDocument(),
    });
    const toolbar = createTableRowToolbar(editor);
    try {
      const s = editor.surface!,
        id = s.session.paragraphIds()[2]!,
        point = { paragraphId: id, offset: 3 };
      s.setSelection({ anchor: point, head: point });
      const anchor = nativeTableAnchor(editor)!,
        table = findNode(s.session.part(), anchor.tableId)!;
      if (table.kind !== 'table') throw new Error('Missing table');
      const rows = table.children.filter((node) => node.kind === 'tableRow');
      const before = canonicalOoxmlFingerprint(s.session.part()),
        kept = rows
          .filter((_, i) => action !== 'delete' || i !== 1)
          .map((row) => [row.id, canonicalOoxmlFingerprint(row)] as const);
      const label = action === 'delete' ? 'Delete row' : `Insert row ${action}`;
      const button = Array.from(toolbar.element.querySelectorAll('button')).find(
        (b) => b.textContent === label
      )!;
      expect(button.disabled).toBe(false);
      const mouse = new MouseEvent('mousedown', { button: 0, bubbles: true, cancelable: true });
      button.dispatchEvent(mouse);
      expect(mouse.defaultPrevented).toBe(true);
      const fixtureId = `table.row.${action === 'above' ? 'before' : action === 'below' ? 'after' : 'delete'}`;
      const fixture = reference.cases.find((c) => c.id === fixtureId)!;
      expect(fixture).toBeDefined();
      const countRows = (value: unknown) =>
        (value as { children: unknown[] }[])[0]!.children.length;
      expect(countRows(fixture.after.value) - countRows(fixture.before.value)).toBe(
        action === 'delete' ? -1 : 1
      );
      expect(changeManualTableRow(editor, action).ok).toBe(true);
      expect(nativeTableCellIds(editor, anchor.tableId)).toHaveLength(action === 'delete' ? 4 : 8);
      for (const [rowId, expected] of kept)
        expect(canonicalOoxmlFingerprint(findNode(s.session.part(), rowId)!)).toEqual(expected);
      expect(nativeTableAnchor(editor)?.tableId).toBe(anchor.tableId);
      const live = findNode(s.session.part(), s.state().selection.head.paragraphId);
      expect(live?.kind).toBe('paragraph');
      if (action !== 'delete') {
        const cells = nativeTableCellIds(editor, anchor.tableId),
          expectedCell = cells[action === 'above' ? 2 : 4]!;
        expect(nativeTableAnchor(editor)?.cellId).toBe(expectedCell);
        expect(paragraphsInCells(s.layout(), [expectedCell])).toContain(
          s.state().selection.head.paragraphId
        );
      }
      const after = canonicalOoxmlFingerprint(s.session.part());
      const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!saved.ok) throw new Error(saved.reason);
      expect(
        canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
      ).toEqual(after);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
    } finally {
      toolbar.dispose();
      editor.destroy();
    }
  });
