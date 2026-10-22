import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { docxFixture, W } from './document';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
export const protectedRun = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
export const protectedParagraph = (text: string) => `<w:p>${protectedRun(text)}</w:p>`;

/** Imported Word content that a body-text writer must retain through save/reopen. */
export function protectedDocument(): Uint8Array {
  const drawing =
    `<w:r><w:drawing xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}">` +
    '<wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Photo" descr="Protected image"/>' +
    `<a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:blipFill><a:blip r:embed="image"/>` +
    '</pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const body =
    '<w:p xmlns:x="urn:opaque" x:keep="paragraph"><w:pPr><w:customKeep w:val="p-property"/></w:pPr>' +
    '<w:r x:keep="run"><w:rPr><w:b/><w:customKeep w:val="r-property"/></w:rPr>' +
    '<w:t x:keep="text" xml:space="preserve">Before </w:t></w:r>' +
    '<w:fldSimple w:instr="PAGE"><w:r><w:t>7</w:t></w:r></w:fldSimple>' +
    protectedRun(' after') +
    '</w:p>' +
    '<w:p>' +
    protectedRun('Left ') +
    '<x:opaque xmlns:x="urn:opaque" x:keep="inline"><x:data>opaque</x:data></x:opaque>' +
    protectedRun('right') +
    '</w:p>' +
    '<x:block xmlns:x="urn:opaque" x:keep="block"><x:data>opaque block</x:data></x:block>' +
    '<w:p>' +
    protectedRun('Complex ') +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    protectedRun('Cached') +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    protectedRun(' tail') +
    '</w:p>' +
    '<w:p>' +
    protectedRun('Photo ') +
    drawing +
    protectedRun(' end') +
    '</w:p>' +
    '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>' +
    protectedRun('Section text') +
    '</w:p>' +
    '<w:sectPr><w:headerReference w:type="default" r:id="header"/>' +
    '<w:footerReference w:type="default" r:id="footer"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/></w:sectPr>';
  const entries = unzipSync(
    docxFixture(body, {
      relationships: ['header', 'footer', 'image', 'footnotes']
        .map(
          (name) =>
            `<Relationship Id="${name}" Type="${R}/${name}" Target="${name === 'image' ? 'media/image.png' : name === 'footnotes' ? 'footnotes.xml' : name + '1.xml'}"/>`
        )
        .join(''),
    })
  );
  entries['word/document.xml'] = strToU8(
    strFromU8(entries['word/document.xml']!)
      .replace('<w:body>', '<w:background w:color="FFFFFF"/><w:body>')
      .replace('Section text</w:t>', 'Section text</w:t><w:footnoteReference w:id="1"/>')
  );
  entries['word/footnotes.xml'] = strToU8(
    `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${protectedParagraph('Protected note')}</w:footnote></w:footnotes>`
  );
  entries['word/header1.xml'] = strToU8(
    `<w:hdr xmlns:w="${W}">${protectedParagraph('Header text')}</w:hdr>`
  );
  entries['word/footer1.xml'] = strToU8(
    `<w:ftr xmlns:w="${W}">${protectedParagraph('Footer text')}</w:ftr>`
  );
  entries['custom/opaque.xml'] = strToU8(
    '<x:data xmlns:x="urn:opaque" x:keep="part">Unknown part</x:data>'
  );
  entries['custom/opaque.bin'] = new Uint8Array([0, 255, 42, 9]);
  entries['word/media/image.png'] = Uint8Array.from(
    atob(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
    ),
    (character) => character.charCodeAt(0)
  );
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Default Extension="png" ContentType="image/png"/>' +
        '<Default Extension="bin" ContentType="application/octet-stream"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'
    )
  );
  return zipSync(entries);
}
