import { docxFixture, W } from './document';
export function fontLinkDocument(character = false) {
  return docxFixture(
    `<w:p><w:r><w:rPr>${character ? '<w:rStyle w:val="Char"/>' : ''}<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:b/></w:rPr><w:t>Alpha</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve"> beta</w:t></w:r></w:p>` +
      '<w:p><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:t>Tail</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
    {
      styles: `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>${character ? '<w:style w:type="character" w:styleId="Base"><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="6.75pt"/><w:szCs w:val="6.75pt"/></w:rPr></w:style><w:style w:type="character" w:styleId="Char"><w:basedOn w:val="Base"/><w:rPr><w:b/></w:rPr></w:style>' : ''}</w:styles>`,
    }
  );
}
export function selectFontLink(
  editor: import('@docx-editor.dev/core/editor').DocxEditorInstance,
  start: number,
  end = start
) {
  const s = editor.surface!,
    id = s.session.paragraphIds()[0]!;
  s.setSelection({
    anchor: { paragraphId: id, offset: start },
    head: { paragraphId: id, offset: end },
  });
}
