import { strToU8, zipSync } from 'fflate';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '@docx-editor.dev/core/store';

export interface BlockData {
  name: string;
  duration?: number;
  answer?: string;
  expanded?: boolean;
}
export interface DocumentBlock {
  id: string;
  tag: string;
  type: string;
  data: BlockData;
  headerId: string;
  headerWordId: string;
  paragraphs: Set<string>;
  children: DocumentBlock[];
}

export function attribute(node: OoxmlElement | undefined, name: string): string {
  return node?.attributes.find((a) => a.localName === name)?.value ?? '';
}
export function elementChild(
  node: OoxmlElement | undefined,
  name: string
): OoxmlElement | undefined {
  for (const c of node?.children ?? [])
    if (c.kind !== 'textValue' && c.localName === name) return c;
  return undefined;
}

export function readBlocks(part: OoxmlPart): DocumentBlock[] {
  const roots: DocumentBlock[] = [];
  const visit = (node: OoxmlNode, parent?: DocumentBlock) => {
    if (node.kind === 'textValue') return;
    let owner = parent;
    if (node.localName === 'sdt') {
      const properties = elementChild(node, 'sdtPr');
      const tag = properties && attribute(elementChild(properties, 'tag'), 'val');
      if (tag?.startsWith('scio:')) {
        const alias = properties && attribute(elementChild(properties, 'alias'), 'val');
        let data: BlockData = { name: tag };
        try {
          const value = JSON.parse(alias || '{}');
          if (typeof value.name === 'string') data = value;
        } catch {
          /* A Word-edited title is not a typed payload. */
        }
        owner = {
          id: node.id,
          tag,
          type: tag.split(':')[1]!,
          data,
          headerId: '',
          headerWordId: '',
          paragraphs: new Set(),
          children: [],
        };
        (parent ? parent.children : roots).push(owner);
      }
    }
    if (node.kind === 'paragraph' && owner) {
      owner.paragraphs.add(node.id);
      if (!owner.headerId) {
        owner.headerId = node.id;
        owner.headerWordId = attribute(node, 'paraId');
      }
    }
    for (const item of node.children) visit(item, owner);
  };
  visit(part.root);
  const includeChildren = (block: DocumentBlock) => {
    for (const nested of block.children) {
      includeChildren(nested);
      for (const id of nested.paragraphs) block.paragraphs.add(id);
    }
  };
  roots.forEach(includeChildren);
  return roots;
}

const escape = (text: string) =>
  text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
let paraId = 0;
const p = (text: string, height?: number) =>
  `<w:p w14:paraId="${(++paraId).toString(16).padStart(8, '0')}"><w:pPr><w:spacing w:before="0" w:after="100"${height ? ` w:line="${height * 20}" w:lineRule="exact"` : ''}/></w:pPr><w:r><w:t xml:space="preserve">${escape(text)}</w:t></w:r></w:p>`;
const sdt = (type: string, data: BlockData, body: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${paraId + 100}"/><w:tag w:val="scio:${type}:one"/><w:alias w:val="${escape(JSON.stringify(data))}"/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;

export function fixture(): Uint8Array {
  paraId = 0;
  const activity = sdt(
    'activity',
    { name: 'Design an experiment', duration: 10 },
    p('Design an experiment — 10 minutes', 62) +
      p('Plan an experiment. Explain your hypothesis and choose the variables.') +
      `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${p('Variable')}</w:tc><w:tc>${p('Observation')}</w:tc></w:tr><w:tr><w:tc>${p('Temperature')}</w:tc><w:tc>${p('Record the result')}</w:tc></w:tr></w:tbl>` +
      Array.from({ length: 5 }, (_, i) =>
        p(
          `Activity step ${i + 1}. Keep these paragraphs as editable Word content inside the React component.`
        )
      ).join('')
  );
  const quiz = sdt(
    'quiz',
    { name: 'Predict the outcome', answer: '', expanded: false },
    p('Predict the outcome. Choose an answer in the browser.', 132)
  );
  const phase = sdt(
    'phase',
    { name: 'Investigation phase' },
    p('Investigation phase', 38) +
      activity +
      quiz +
      p('Reflection inside the phase. Compare the prediction with your observations.')
  );
  const body =
    p('React blocks in a Word document') +
    p('Ordinary editable text before the custom blocks.') +
    phase +
    p('Ordinary editable text after the custom blocks.') +
    Array.from({ length: 5 }, (_, i) =>
      p(`Closing paragraph ${i + 1}. Pagination follows the native document tree.`)
    ).join('') +
    '<w:sectPr><w:pgSz w:w="8400" w:h="9600"/><w:pgMar w:top="480" w:bottom="480" w:left="480" w:right="480"/></w:sectPr>';
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}</w:body></w:document>`
    ),
  });
}
