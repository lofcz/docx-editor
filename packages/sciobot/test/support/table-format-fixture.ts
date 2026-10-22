import {
  bodyStoryRoot,
  findNode,
  paragraphTextOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { namedChild, nodeChildren, rowCells, tableRows } from './native-table-oracle';

export const formatParagraph = (text: string) =>
  `<w:p><w:pPr><keep:extra xmlns:keep="urn:keep" keep:value="paragraph"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
export const formatCell = (text: string, width = 2400, extra = '') =>
  `<w:tc><w:tcPr xmlns:keep="urn:keep"><w:tcW w:w="${width}" w:type="dxa" keep:value="cell-width"><keep:child/></w:tcW><w:tcBorders><w:top w:val="single" w:sz="4" w:color="111111" keep:value="top"/><w:left w:val="double" w:sz="8" w:color="222222"/><w:bottom w:val="dashed" w:sz="12" w:color="333333"/><w:right w:val="single" w:sz="16" w:color="444444"/></w:tcBorders><w:shd w:val="pct10" w:color="112233" w:fill="EEEEEE" keep:value="shading"/><w:vAlign w:val="top" keep:value="alignment"/><keep:extra keep:value="cell"/>${extra}</w:tcPr>${formatParagraph(text)}</w:tc>`;
export const formatRow = (a: string, b: string, c: string, header = false) =>
  `<w:tr><w:trPr xmlns:keep="urn:keep"><w:cantSplit/><w:trHeight w:val="400" w:hRule="atLeast" keep:value="height"><keep:child/></w:trHeight>${header ? '<w:tblHeader w:val="1"/>' : ''}</w:trPr>${formatCell(a, 1200)}${formatCell(b)}${formatCell(c, 3600)}</w:tr>`;
export const formatTable = (rows: string) =>
  `<w:tbl><w:tblPr xmlns:keep="urn:keep"><w:tblStyle w:val="Original"/><w:tblW w:w="7200" w:type="dxa" keep:value="table-width"><keep:child/></w:tblW><w:tblLayout w:type="fixed" keep:value="layout"><keep:child/></w:tblLayout><w:tblCaption w:val="Keep caption"/><keep:extra keep:value="table"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>${rows}</w:tbl>`;
export const formatBody = (merged = false) =>
  formatParagraph('Before') +
  formatTable(
    merged
      ? '<w:tr>' +
          formatCell('AB1', 3600, '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') +
          formatCell('C1', 3600) +
          '</w:tr><w:tr>' +
          formatCell('', 3600, '<w:gridSpan w:val="2"/><w:vMerge/>') +
          formatCell('C2', 3600) +
          '</w:tr>' +
          formatRow('A3', 'B3', 'C3')
      : formatRow('A1', 'B1', 'C1', true) +
          formatRow('A2', 'B2', 'C2') +
          formatRow('A3', 'B3', 'C3')
  ) +
  formatParagraph('Between') +
  formatTable(formatRow('Other A', 'Other B', 'Other C')) +
  formatParagraph('After');
export const formatTables = (part: OoxmlPart) =>
  nodeChildren(bodyStoryRoot(part)!).filter((node) => node.kind === 'table');
export const formatAttr = (node: OoxmlNode, name: string, namespaceUri = WML_NAMESPACE_URI) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find(
        (entry) => entry.namespaceUri === namespaceUri && entry.localName === name
      )?.value;
export function formatProperty(node: OoxmlNode, container: string, property: string): OoxmlNode {
  const result = namedChild(namedChild(node, container)!, property);
  if (!result) throw new Error(`Missing property ${container}/${property}`);
  return result;
}
export function formatDescendants(node: OoxmlNode): OoxmlNode[] {
  return [node, ...nodeChildren(node).flatMap(formatDescendants)];
}
export const formatContent = (part: OoxmlPart, tableId: string) =>
  tableRows(part, tableId).map((row) =>
    rowCells(row).map((cell) =>
      formatDescendants(cell)
        .filter((node) => node.kind === 'paragraph')
        .map((node) => ({ id: node.id, text: paragraphTextOf(part, node.id) }))
    )
  );
export const formatCells = (part: OoxmlPart, tableId: string) =>
  tableRows(part, tableId).map(rowCells);
export const formatNode = (part: OoxmlPart, id: string) => {
  const node = findNode(part, id);
  if (!node) throw new Error(`Missing node ${id}`);
  return node;
};
