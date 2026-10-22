import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { docxFixture, W } from './document';

export const listPPr = (numId: string, level = 0) =>
  `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`;
export const listParagraph = (text: string, numId?: string, level = 0) =>
  `<w:p><w:pPr>${numId ? listPPr(numId, level) : ''}<w:spacing w:after="120"/><w:customKeep w:val="yes"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
export const checklistParagraph = (checked: boolean, level = 0, tag = 'scio:list-checkbox') =>
  `<w:p><w:pPr><w:ind w:left="${720 * (level + 1)}" w:hanging="360"/><w:customKeep w:val="yes"/></w:pPr>` +
  `<w:sdt><w:sdtPr><w:id w:val="8"/><w:tag w:val="${tag}"/><w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checked w14:val="${checked ? 1 : 0}"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>${checked ? '☒' : '☐'}</w:t></w:r></w:sdtContent></w:sdt>` +
  '<w:r><w:t>Task</w:t></w:r></w:p>';
export const listNumbering =
  `<w:numbering xmlns:w="${W}">` +
  ['decimal', 'bullet', 'lowerRoman']
    .map(
      (format, index) =>
        `<w:abstractNum w:abstractNumId="${index}">` +
        Array.from(
          { length: 9 },
          (_, level) =>
            `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${format === 'bullet' ? '•' : '%' + (level + 1) + '.'}"/><w:pPr><w:ind w:left="${720 * (level + 1)}" w:hanging="360"/></w:pPr></w:lvl>`
        ).join('') +
        '</w:abstractNum>'
    )
    .join('') +
  [0, 1, 2]
    .map((index) => `<w:num w:numId="${index + 1}"><w:abstractNumId w:val="${index}"/></w:num>`)
    .join('') +
  '</w:numbering>';

export function listDocument(body: string, styles?: string): Uint8Array {
  const bytes = docxFixture(body, {
    numbering: listNumbering,
    ...(styles
      ? {
          relationships:
            '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
        }
      : {}),
  });
  if (!styles) return bytes;
  const entries = unzipSync(bytes);
  entries['word/styles.xml'] = strToU8(styles);
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
    )
  );
  return zipSync(entries);
}
