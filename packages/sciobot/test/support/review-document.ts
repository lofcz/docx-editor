import { docxFixture } from './document';

/** Native insert, delete, run-format, and paragraph-format revisions in one real DOCX. */
export function reviewDocumentFixture() {
  return docxFixture(
    '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="Ada"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t xml:space="preserve">Kept </w:t></w:r>' +
      '<w:del w:id="2" w:author="Ada"><w:r><w:delText>Old</w:delText></w:r></w:del>' +
      '<w:ins w:id="3" w:author="Ada"><w:r><w:t>New</w:t></w:r></w:ins></w:p>' +
      '<w:p><w:pPr><w:outlineLvl w:val="0"/><w:pPrChange w:id="4" w:author="Ada"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>'
  );
}
