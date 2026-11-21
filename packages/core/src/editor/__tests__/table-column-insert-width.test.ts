import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { canonicalOoxmlFingerprint } from '../../store/package/ooxml-tree.ts';
import { readEditableTableTopology } from '../../store/store/tree-op-table-topology.ts';
import { wmlAttributeValue, wmlChildNamed } from '../../store/store/tree-op-table-shared.ts';
import { paragraphTextOf } from '../../store/store/tree-ops.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { formatDefaultsDocx, paragraph } from './paginated-surface-fixtures.ts';

function documentWithTable(properties: string, style?: string) {
  const widths = [1800, 3000, 4060];
  const row = (values: number[]) =>
    '<w:tr>' +
    values
      .map(
        (w, i) =>
          `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${w}"/></w:tcPr>${paragraph(String(i + 1))}</w:tc>`
      )
      .join('') +
    '</w:tr>';
  const data = formatDefaultsDocx(
    `<w:tbl><w:tblPr>${properties}</w:tblPr><w:tblGrid>` +
      widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('') +
      '</w:tblGrid>' +
      row(widths) +
      row([1650, 2700, 3600]) +
      '</w:tbl><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1523" w:right="1523"/></w:sectPr>'
  );
  if (!style) return data;
  const parts = unzipSync(data);
  parts['word/styles.xml'] = strToU8(
    strFromU8(parts['word/styles.xml']!).replace('</w:styles>', style + '</w:styles>')
  );
  return zipSync(parts);
}

function read(editor: DocxEditorInstance) {
  const surface = editor.surface!;
  const part = surface.session.part();
  const fragment = surface.layout().pages[0]!.fragments.find((f) => f.kind === 'table');
  if (fragment?.kind !== 'table') throw new Error('missing table');
  const result = readEditableTableTopology(part.root, fragment.tableId);
  if (!result.ok) throw new Error(result.reason);
  const { table, rows } = result.topology;
  const props = wmlChildNamed(table, 'tblPr')!;
  const width = wmlChildNamed(props, 'tblW');
  const layout = wmlChildNamed(props, 'tblLayout');
  return {
    width: Math.round(fragment.box.width * 20),
    preferred: width && [wmlAttributeValue(width, 'type'), wmlAttributeValue(width, 'w')],
    layout: layout && wmlAttributeValue(layout, 'type'),
    cells: rows.map(({ cells }) =>
      cells.map((cell) => {
        const p = cell.children.find((n) => n.kind === 'paragraph')!;
        const w = wmlChildNamed(wmlChildNamed(cell, 'tcPr')!, 'tcW')!;
        return { text: paragraphTextOf(part, p.id), width: wmlAttributeValue(w, 'w') };
      })
    ),
  };
}

const fixed = '<w:tblLayout w:type="fixed"/>';
const auto = '<w:tblW w:type="auto" w:w="0"/>';
const percent = '<w:tblW w:type="pct" w:w="5000"/>';
const oversized = '<w:tblW w:type="dxa" w:w="15000"/>';
const inherited =
  `<w:style w:type="table" w:styleId="Base"><w:tblPr>${percent}${fixed}</w:tblPr></w:style>` +
  '<w:style w:type="table" w:styleId="Derived"><w:basedOn w:val="Base"/></w:style>';

for (const example of [
  {
    name: 'fixed layout with automatic width',
    properties: auto + fixed,
    width: 8860,
    preferred: ['dxa', '8860'],
  },
  {
    name: 'AutoFit with automatic width',
    properties: auto,
    width: 8860,
    preferred: ['dxa', '8860'],
  },
  {
    name: 'percentage width',
    properties: percent + fixed,
    width: 8860,
    preferred: ['pct', '5000'],
  },
  {
    name: 'explicit oversized width',
    properties: oversized + fixed,
    width: 15000,
    preferred: ['dxa', '15000'],
  },
  {
    name: 'inherited percentage width',
    properties: '<w:tblStyle w:val="Derived"/>',
    style: inherited,
    width: 8860,
    preferred: undefined,
  },
]) {
  test(`column insertion retains table size and cell preferences for ${example.name}`, async () => {
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: documentWithTable(example.properties, example.style),
    });
    try {
      const surface = editor.surface!;
      const position = { paragraphId: surface.session.paragraphIds()[2]!, offset: 0 };
      surface.setSelection({ head: position, anchor: position });
      const before = read(editor);
      expect(before.width).toBe(example.width);
      const sheet = container.querySelector<HTMLElement>('.docx-page')!;
      const extendsPastPaper = example.width === 15000;
      expect(sheet.hasAttribute('data-table-overflow')).toBe(extendsPastPaper);
      if (extendsPastPaper) {
        expect(parseFloat(sheet.style.getPropertyValue('--docx-table-clip-right'))).toBeLessThan(0);
        const pages = container.querySelector<HTMLElement>('.docx-pages')!;
        expect(parseFloat(pages.style.width)).toBeGreaterThan(parseFloat(sheet.style.width));
      }
      const fingerprint = canonicalOoxmlFingerprint(surface.session.part());
      expect(editor.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
      const first = read(editor);
      expect(first.width).toBe(example.width);
      expect(first.preferred).toEqual(example.preferred);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(fingerprint);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(read(editor)).toEqual(first);
      expect(editor.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
      const after = read(editor);
      expect(after.width).toBe(example.width);
      expect(after.layout).toBe(before.layout);
      expect(after.preferred).toEqual(example.preferred);
      for (const [index, row] of after.cells.entries()) {
        expect(row.slice(0, 3)).toEqual(before.cells[index]!);
        expect(row.slice(3)).toEqual(Array(2).fill({ text: '', width: row[2]!.width }));
      }
      await editor.load(await editor.save());
      expect(read(editor)).toEqual(after);
    } finally {
      editor.destroy();
      container.remove();
    }
  });
}
