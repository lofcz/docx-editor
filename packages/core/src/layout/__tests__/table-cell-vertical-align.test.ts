import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function part(xml: string, name = '/word/document.xml') {
  const result = readOoxmlPart(xml, { name, contentType: 'app/xml' });
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
const style = (id: string, properties: string, base = '') =>
  `<w:style w:type="table" w:styleId="${id}">${base ? `<w:basedOn w:val="${base}"/>` : ''}${properties}</w:style>`;
function render(styles: string, direct = '') {
  const document = part(
    `<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="Child"/><w:tblLook w:firstRow="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="2000" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr>${direct}</w:tcPr><w:p><w:r><w:t>Aligned</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`
  );
  const source = serializeOoxmlPart(document);
  const styleCascade = buildStyleCascadeTable(
    part(`<w:styles xmlns:w="${W}">${styles}</w:styles>`, '/word/styles.xml').root
  );
  const options = { measurer: createFixedMeasurer(), styleCascade };
  const layout = layoutSemanticDocument(document, 0, options);
  const table = layout.pages[0]!.fragments[0]!;
  if (table.kind !== 'table') throw Error('Missing table');
  const cell = table.rows[0]!.cells[0]!;
  const reopened = part(source);
  expect(layoutSemanticDocument(reopened, 0, options)).toEqual(layout);
  expect(serializeOoxmlPart(document)).toBe(source);
  return cell;
}

const inherited =
  style('Base', '<w:tcPr><w:vAlign w:val="bottom"/></w:tcPr>') + style('Child', '', 'Base');
test('table basedOn alignment controls the effective value and content position', () => {
  const cell = render(inherited);
  expect(cell.vAlign).toBe('bottom');
  expect(cell.blocks[0]!.box.y).toBeGreaterThan(cell.box.y + 60);
});
test('a direct top value overrides inherited bottom alignment', () => {
  const cell = render(inherited, '<w:vAlign w:val="top"/>');
  expect(cell.vAlign).toBe('top');
  expect(cell.blocks[0]!.box.y).toBeLessThan(cell.box.y + 5);
});
test('conditional alignment survives an inherited condition layer that omits it', () => {
  const styles =
    style(
      'Base',
      '<w:tcPr><w:vAlign w:val="bottom"/></w:tcPr><w:tblStylePr w:type="firstRow"><w:tcPr><w:vAlign w:val="center"/></w:tcPr></w:tblStylePr>'
    ) +
    style(
      'Child',
      '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr></w:tblStylePr>',
      'Base'
    );
  const cell = render(styles);
  expect(cell.vAlign).toBe('center');
  expect(cell.blocks[0]!.box.y).toBeGreaterThan(cell.box.y + 25);
  expect(cell.blocks[0]!.box.y).toBeLessThan(cell.box.y + 60);
});
test('direct alignment overrides conditional style alignment', () => {
  const styles = style(
    'Child',
    '<w:tblStylePr w:type="firstRow"><w:tcPr><w:vAlign w:val="center"/></w:tcPr></w:tblStylePr>'
  );
  expect(render(styles, '<w:vAlign w:val="bottom"/>').vAlign).toBe('bottom');
});
