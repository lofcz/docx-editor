import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, findNode, paragraphTextOf } from '@docx-editor.dev/core/store';
import { manualColumnDocument, columnShape } from '../../test/support/manual-table-columns';
import { savedTableFingerprint } from '../../test/support/manual-table-controls';
import {
  changeManualTableAppearance,
  createTableAppearanceToolbar,
  type TableAppearanceChange,
} from './table-appearance';
import { namedChild, nodeChildren } from '../../test/support/native-table-oracle';
const changes: TableAppearanceChange[] = [
  { kind: 'fill', color: '1A2B3C' },
  { kind: 'column-width', twips: 2100 },
  { kind: 'row-height', twips: 720 },
  { kind: 'borders', color: '123456', size: 8 },
  { kind: 'headers', count: 2 },
];
for (const merge of ['none', 'horizontal', 'vertical'] as const)
  test(`table appearance preserves ${merge} merges, neighbors, imported properties and exact per-gesture history`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: manualColumnDocument(merge),
    });
    const toolbar = createTableAppearanceToolbar(editor);
    try {
      const s = editor.surface!,
        source = columnShape(s.session.part()),
        cell = source.rows[1]!.cells[1]!,
        point = { paragraphId: cell.paragraphs[0]!.id, offset: 2 };
      s.setSelection({ anchor: point, head: point });
      const original = s.session.part(),
        snapshots = [canonicalOoxmlFingerprint(original)];
      for (const change of changes) {
        expect(changeManualTableAppearance(editor, change).ok).toBe(true);
        snapshots.push(canonicalOoxmlFingerprint(s.session.part()));
      }
      const after = columnShape(s.session.part());
      expect(after.widths).toEqual([1200, 2100, 2400]);
      expect(after.rows.map((r) => r.header)).toEqual([true, true, false]);
      for (const row of source.rows)
        for (const c of row.cells)
          for (const p of c.paragraphs) {
            expect(paragraphTextOf(s.session.part(), p.id)).toBe(p.text);
            expect(canonicalOoxmlFingerprint(findNode(s.session.part(), p.id)!)).toEqual(
              canonicalOoxmlFingerprint(findNode(original, p.id)!)
            );
          }
      const current = findNode(s.session.part(), cell.id)!,
        props = namedChild(current, 'tcPr')!,
        fill = namedChild(props, 'shd')!,
        borders = namedChild(props, 'tcBorders')!;
      expect(
        fill.kind === 'textValue'
          ? null
          : fill.attributes.find((a) => a.localName === 'fill')?.value
      ).toBe('1A2B3C');
      expect(nodeChildren(borders)).toHaveLength(4);
      for (const edge of nodeChildren(borders)) {
        if (edge.kind === 'textValue') throw new Error('Invalid border');
        const values = Object.fromEntries(edge.attributes.map((a) => [a.localName, a.value]));
        expect(values).toMatchObject({ val: 'single', sz: '8', color: '123456' });
      }
      const row = findNode(s.session.part(), source.rows[1]!.id)!;
      const height = namedChild(namedChild(row, 'trPr')!, 'trHeight')!;
      if (height.kind === 'textValue') throw new Error('Invalid row height');
      expect(
        Object.fromEntries(height.attributes.map((a) => [a.localName, a.value]))
      ).toMatchObject({ val: '720', hRule: 'exact' });
      expect(await savedTableFingerprint(editor)).toEqual(snapshots.at(-1)!);
      for (let i = changes.length - 1; i >= 0; i--) {
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(snapshots[i]!);
      }
      for (let i = 1; i < snapshots.length; i++) {
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(snapshots[i]!);
      }
      expect(changeManualTableAppearance(editor, { kind: 'fill', color: null }).ok).toBe(true);
      expect(changeManualTableAppearance(editor, { kind: 'clear-borders' }).ok).toBe(true);
    } finally {
      toolbar.dispose();
      editor.destroy();
    }
  });
test('invalid appearance values and a viewing editor perform no writes', () => {
  for (const mode of ['edit', 'view'] as const) {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: manualColumnDocument('none'),
      mode,
    });
    try {
      const s = editor.surface!,
        shape = columnShape(s.session.part()),
        point = { paragraphId: shape.rows[1]!.cells[1]!.paragraphs[0]!.id, offset: 0 };
      s.setSelection({ anchor: point, head: point });
      const before = canonicalOoxmlFingerprint(s.session.part());
      for (const change of [
        { kind: 'column-width', twips: -1 },
        { kind: 'row-height', twips: NaN },
        { kind: 'fill', color: 'url(bad)' },
        { kind: 'headers', count: 4 },
      ] as TableAppearanceChange[])
        expect(changeManualTableAppearance(editor, change).ok).toBe(false);
      if (mode === 'view') {
        const locked = changeManualTableAppearance(editor, { kind: 'fill', color: '123456' });
        expect(locked.ok).toBe(false);
        if (!locked.ok) expect(locked.code).toBe('locked');
      }
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    } finally {
      editor.destroy();
    }
  }
});
