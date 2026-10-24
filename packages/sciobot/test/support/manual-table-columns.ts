import {
  paragraphTextOf,
  type OoxmlPart,
  type OoxmlNode,
  type OoxmlElement,
} from '@docx-editor.dev/core/store';
import { docxFixture } from './document';

export type ColumnMerge = 'horizontal' | 'vertical' | 'none';
/** Nonuniform imported widths, a merged neighbor, header behavior and foreign metadata. */
export function manualColumnDocument(merge: ColumnMerge, grid = true) {
  const cell = (text: string, width: number, properties = '') =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${properties}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const rows = ['Header', 'Middle', 'Last'].map((label, index) => {
    const merged = merge === 'horizontal' && index === 0;
    const cells = merged
      ? cell('Header AB', 3000, '<w:gridSpan w:val="2"/>') + cell('Header C', 2400)
      : cell(
          `${label} A`,
          1200,
          merge === 'vertical' ? `<w:vMerge${index === 0 ? ' w:val="restart"' : ''}/>` : ''
        ) +
        cell(`${label} B`, 1800) +
        cell(`${label} C`, 2400);
    return `<w:tr xmlns:keep="urn:keep" keep:row="${label}"><w:trPr>${index === 0 ? '<w:tblHeader/>' : ''}<w:cantSplit/></w:trPr>${cells}</w:tr>`;
  });
  const columns = grid
    ? '<w:tblGrid>' +
      [1200, 1800, 2400]
        .map((width, i) => `<w:gridCol xmlns:keep="urn:keep" keep:column="${i}" w:w="${width}"/>`)
        .join('') +
      '</w:tblGrid>'
    : '';
  return docxFixture(
    `<w:tbl><w:tblPr/>${columns}${rows.join('')}</w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p>`
  );
}

function elements(node: OoxmlElement) {
  const nodes: readonly OoxmlNode[] = node.children;
  return nodes.filter((n): n is OoxmlElement => n.kind !== 'textValue');
}
const child = (node: OoxmlElement, name: string) =>
  elements(node).find(
    (n) =>
      n.namespaceUri === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' &&
      n.localName === name
  );
const value = (node: OoxmlElement | undefined, name: string) =>
  node?.attributes.find(
    (a) =>
      a.namespaceUri === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' &&
      a.localName === name
  )?.value;

/** Independent XML-property oracle; it does not use the native editing planner or grid placement. */
export function columnShape(part: OoxmlPart) {
  const body = child(part.root, 'body')!,
    table = child(body, 'tbl')!,
    grid = child(table, 'tblGrid');
  return {
    widths: (grid ? elements(grid) : [])
      .filter((n) => n.localName === 'gridCol')
      .map((n) => Number(value(n, 'w'))),
    rows: elements(table)
      .filter((n) => n.kind === 'tableRow')
      .map((row) => ({
        id: row.id,
        header: !!child(child(row, 'trPr')!, 'tblHeader'),
        cells: elements(row)
          .filter((n) => n.kind === 'tableCell')
          .map((cell) => {
            const properties = child(cell, 'tcPr')!,
              vm = child(properties, 'vMerge');
            return {
              id: cell.id,
              width: Number(value(child(properties, 'tcW'), 'w')),
              span: Number(value(child(properties, 'gridSpan'), 'val') ?? 1),
              vMerge: vm ? (value(vm, 'val') ?? 'continue') : null,
              paragraphs: elements(cell)
                .filter((n) => n.kind === 'paragraph')
                .map((p) => ({ id: p.id, text: paragraphTextOf(part, p.id) })),
            };
          }),
      })),
  };
}
