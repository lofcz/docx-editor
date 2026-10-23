import { escapeXml, readOoxmlPart, type OoxmlNode } from '@docx-editor.dev/core/store';
type OoxmlDrawingNode = Extract<OoxmlNode, { kind: 'drawing' }>;
/** Application-owned Word picture constructor over the public Apache XML reader. */
export function inlinePicture(input: {
  readonly docPrId: number;
  readonly relationshipId: string;
  readonly widthPoints: number;
  readonly heightPoints: number;
  readonly title: string | null;
  readonly alt: string;
}): OoxmlDrawingNode {
  const cx = Math.round(input.widthPoints * 12700),
    cy = Math.round(input.heightPoints * 12700);
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${input.docPrId}" name="" descr="${escapeXml(input.alt)}"${input.title !== null ? ` title="${escapeXml(input.title)}"` : ''}/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${escapeXml(input.relationshipId)}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`;
  const parsed = readOoxmlPart(xml, {
    name: '/sciobot-inline-image.xml',
    contentType: 'application/xml',
  });
  if (!parsed.ok) throw Error(`Native picture XML refused: ${parsed.reason}.`);
  let drawing: OoxmlDrawingNode | undefined;
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'drawing') drawing = node;
    else if (node.kind !== 'textValue') node.children.forEach(visit);
  };
  visit(parsed.part.root);
  if (!drawing) throw Error('The picture XML has no native drawing.');
  return drawing;
}
