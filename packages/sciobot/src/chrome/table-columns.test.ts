import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, findNode, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { manualColumnDocument, columnShape } from '../../test/support/manual-table-columns';
import { nativeTableAnchor } from './tables';
import { createTableColumnToolbar } from './table-columns';

for (const merge of ['horizontal', 'vertical'] as const)
  for (const action of ['left', 'right', 'delete'] as const)
    test(`column ${action} beside ${merge} merges retains widths, spans, contents, exact history and Word output`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: manualColumnDocument(merge),
      });
      const toolbar = createTableColumnToolbar(editor);
      try {
        const s = editor.surface!,
          source = columnShape(s.session.part());
        const point = { paragraphId: source.rows[1]!.cells[1]!.paragraphs[0]!.id, offset: 3 };
        s.setSelection({ anchor: point, head: point });
        const beforePart = s.session.part();
        const before = canonicalOoxmlFingerprint(beforePart);
        const label = action === 'delete' ? 'Delete column' : `Insert column ${action}`;
        const button = Array.from(toolbar.element.querySelectorAll('button')).find(
          (b) => b.textContent === label
        )!;
        expect(button.disabled).toBe(false);
        const mouse = new MouseEvent('mousedown', { button: 0, bubbles: true, cancelable: true });
        button.dispatchEvent(mouse);
        expect(mouse.defaultPrevented).toBe(true);
        button.click();
        const changed = columnShape(s.session.part());
        expect(changed.widths).toEqual(
          action === 'delete' ? [1200, 2400] : [1200, 1800, 1800, 2400]
        );
        expect(changed.rows.map((r) => r.header)).toEqual([true, false, false]);
        for (const row of changed.rows)
          expect(row.cells.reduce((n, c) => n + c.span, 0)).toBe(changed.widths.length);
        if (merge === 'horizontal') {
          const cell = changed.rows[0]!.cells[0]!;
          expect(cell.span).toBe(action === 'left' ? 3 : action === 'right' ? 2 : 1);
          expect(cell.width).toBe(action === 'left' ? 4800 : action === 'right' ? 3000 : 1200);
          expect(cell.paragraphs[0]!.text).toBe('Header AB');
        } else
          expect(changed.rows.map((r) => r.cells[0]!.vMerge)).toEqual([
            'restart',
            'continue',
            'continue',
          ]);
        for (const row of source.rows)
          for (const cell of row.cells)
            for (const paragraph of cell.paragraphs) {
              const original = findNode(s.session.part(), paragraph.id);
              if (
                action !== 'delete' ||
                !row.cells[1]!.paragraphs.some((p) => p.id === paragraph.id) ||
                (merge === 'horizontal' && row === source.rows[0])
              ) {
                expect(original?.kind).toBe('paragraph');
                expect(canonicalOoxmlFingerprint(original!)).toEqual(
                  canonicalOoxmlFingerprint(findNode(beforePart, paragraph.id)!)
                );
              }
            }
        expect(nativeTableAnchor(editor)).not.toBeNull();
        expect(findNode(s.session.part(), s.state().selection.head.paragraphId)?.kind).toBe(
          'paragraph'
        );
        if (action !== 'delete') {
          const live = changed.rows
            .flatMap((r) => r.cells)
            .flatMap((c) => c.paragraphs)
            .find((p) => p.id === s.state().selection.head.paragraphId)!;
          expect(live.text).toBe('');
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
