import { findNode, readOoxmlPackage, type OoxmlNode } from '@docx-editor.dev/core/store';
import { docxFixture } from './document';

/** Explicit Word source oracle, independent of the table property operation. */
export function tableIndentDocument(twips?: number, merged = false) {
  const cell = (text: string, width: number, extra = '') =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  let paragraph = 0;
  return docxFixture(
    `<w:tbl xmlns:k="urn:keep" k:table="metadata"><w:tblPr><w:tblW w:w="3600" w:type="dxa"/><w:jc w:val="left"/>${twips === undefined ? '' : `<w:tblInd w:w="${twips}" w:type="dxa" k:tag="indent"><k:inside/></w:tblInd>`}<w:tblLayout w:type="fixed"/><w:tblCaption w:val="Imported caption"/><k:extra/></w:tblPr><w:tblGrid><w:gridCol w:w="1200" k:slot="a"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>${merged ? cell('Header AB', 3600, '<w:gridSpan w:val="2"/>') : cell('Header A', 1200) + cell('Header B', 2400)}</w:tr><w:tr k:row="keep">${cell('Middle A', 1200) + cell('Middle B', 2400)}</w:tr></w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p>`.replaceAll(
      '<w:p>',
      () => {
        const id = (++paragraph).toString(16).padStart(8, '0').toUpperCase();
        return `<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="${id}" w14:textId="${id}">`;
      }
    )
  );
}
export function indentPart(twips?: number, merged = false) {
  const read = readOoxmlPackage(tableIndentDocument(twips, merged));
  if (!read.ok) throw new Error(read.reason);
  return read.package.parts.get(read.package.mainDocumentPart)!;
}
export function indentTable(part: ReturnType<typeof indentPart>) {
  const body = part.root.children.find((n) => n.kind !== 'textValue' && n.localName === 'body');
  if (!body || body.kind === 'textValue') throw new Error('Missing body');
  return body.children.find((n) => n.kind === 'table')!;
}
export function indentLeaf(part: ReturnType<typeof indentPart>) {
  const table = findNode(part, indentTable(part).id)!;
  if (table.kind === 'textValue') throw new Error('Missing table');
  const properties = table.children.find((n) => n.kind === 'tableProperties');
  return properties && properties.kind !== 'textValue'
    ? properties.children.find((n) => n.kind !== 'textValue' && n.localName === 'tblInd')
    : undefined;
}
export function indentValue(
  node: OoxmlNode | undefined,
  name: string,
  namespaceUri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
) {
  return node && node.kind !== 'textValue'
    ? node.attributes.find((a) => a.namespaceUri === namespaceUri && a.localName === name)?.value
    : undefined;
}
