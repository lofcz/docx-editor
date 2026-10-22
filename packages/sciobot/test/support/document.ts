import { readOoxmlPart } from '@docx-editor.dev/core/store';
import { strToU8, zipSync } from 'fflate';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function bodyPart(body: string) {
  const loaded = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!loaded.ok) throw new Error(`Invalid test document: ${loaded.reason}`);
  return loaded.part;
}

/** Minimal real OPC document for tests of native import, layout, and projection. */
export function docxFixture(
  body: string,
  options: { numbering?: string; relationships?: string; styles?: string } = {}
): Uint8Array {
  const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const numberingType =
    'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
  const stylesType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        (options.numbering
          ? `<Override PartName="/word/numbering.xml" ContentType="${numberingType}"/>`
          : '') +
        (options.styles
          ? `<Override PartName="/word/styles.xml" ContentType="${stylesType}"/>`
          : '') +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${rel}"><Relationship Id="document" Type="${relationship}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${relationship}"><w:body>${body}</w:body></w:document>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${rel}">${options.numbering ? `<Relationship Id="numbering" Type="${relationship}/numbering" Target="numbering.xml"/>` : ''}${options.styles ? `<Relationship Id="styles" Type="${relationship}/styles" Target="styles.xml"/>` : ''}${options.relationships ?? ''}</Relationships>`
    ),
  };
  if (options.numbering) entries['word/numbering.xml'] = strToU8(options.numbering);
  if (options.styles) entries['word/styles.xml'] = strToU8(options.styles);
  return zipSync(entries);
}
