import type { TreeDocOp } from '@docx-editor.dev/core/store';
const date = '2026-10-24T12:00:00Z';
export const propertyBoundaryBody =
  '<w:p xmlns:u="urn:opaque"><w:pPr><w:jc w:val="left"/><u:rPr u:owned="yes"><u:opaque/></u:rPr>' +
  '<w:rPr><w:ins w:id="77" w:author="Peer"/><w:b w:val="0"/></w:rPr>' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>';
export function propertyBoundaryOps(id: string): TreeDocOp[] {
  return [
    {
      op: 'setParagraphProperties',
      paragraphId: id,
      properties: [{ localName: 'jc', attributes: { val: 'center' } }],
      revision: { author: 'Alice', date },
    },
    {
      op: 'setParagraphMarkProperties',
      paragraphId: id,
      properties: [{ localName: 'b', attributes: {} }],
      revision: { author: 'Bot', date },
    },
  ];
}
