import {
  canonicalOoxmlFingerprint,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from './document';

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
/** Imported fraction with retained properties, Unicode, and a vendor extension. */
export function mathDocument(): Uint8Array {
  return docxFixture(
    '<w:p><w:r><w:t>Before</w:t></w:r>' +
      `<m:oMath xmlns:m="${M}" xmlns:v="urn:test:math"><m:f><m:fPr><m:type m:val="bar"/><v:keep v:flag="original"/></m:fPr>` +
      '<m:num><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>α+1</m:t></m:r></m:num>' +
      '<m:den><m:r><m:t>β</m:t></m:r></m:den></m:f></m:oMath>' +
      '<w:r><w:t>After</w:t></w:r></w:p>' +
      `<w:p><m:oMathPara xmlns:m="${M}"><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>` +
      '<m:oMath><m:r><m:t>display = 𝛼</m:t></m:r></m:oMath></m:oMathPara></w:p>' +
      '<w:p><w:r><w:t>Neighbor</w:t></w:r></w:p>'
  );
}
export function mathNodes(root: OoxmlNode): OoxmlNode[] {
  if (root.kind === 'textValue') return [];
  if (root.namespaceUri === M && (root.localName === 'oMath' || root.localName === 'oMathPara'))
    return [root];
  return root.children.flatMap(mathNodes);
}
export function importedMath(pkg: OoxmlPackage): string[] {
  return mathNodes(pkg.parts.get(pkg.mainDocumentPart)!.root).map(canonicalOoxmlFingerprint);
}
