import { expect, test } from 'bun:test';
import { exportMarkdownFrom } from '../src/index.ts';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import { docx } from './fixture.ts';

test('nested restarts keep their parent indentation within the same numbering definition', async () => {
  const level = (depth: number, format: string, label: string) =>
    `<w:lvl w:ilvl="${depth}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${label}"/></w:lvl>`;
  const levels = level(0, 'bullet', '•') + level(1, 'decimal', '%2.');
  const instance = (id: number, definition: number, start?: number) =>
    `<w:num w:numId="${id}"><w:abstractNumId w:val="${definition}"/>` +
    (start ? `<w:lvlOverride w:ilvl="1"><w:startOverride w:val="${start}"/></w:lvlOverride>` : '') +
    '</w:num>';
  const numbering =
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:abstractNum w:abstractNumId="0">${levels}</w:abstractNum><w:abstractNum w:abstractNumId="1">${levels}</w:abstractNum>` +
    instance(1, 0) +
    instance(2, 0, 5) +
    instance(3, 0, 9) +
    instance(4, 1, 3) +
    '</w:numbering>';
  const item = (id: number, depth: number, value: string) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${depth}"/><w:numId w:val="${id}"/></w:numPr></w:pPr><w:r><w:t>${value}</w:t></w:r></w:p>`;
  const body =
    item(1, 0, 'Parent') +
    item(2, 1, 'Restart') +
    item(2, 1, 'Sibling') +
    item(1, 0, 'Next parent') +
    item(3, 1, 'Second restart') +
    item(4, 1, 'Independent') +
    '<w:p><w:r><w:t>Break</w:t></w:r></w:p>' +
    item(3, 1, 'Orphan');
  const opened = openDocumentForExport(docx(body, numbering));
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const result = await exportMarkdownFrom(opened.session);
    const expected = [
      '- Parent',
      '  5. Restart',
      '  6. Sibling',
      '- Next parent',
      '  9. Second restart',
      '3. Independent',
      'Break',
      '10. Orphan',
    ].join('\n\n');
    expect(result.markdown).toBe(expected);
    expect(result.pages[0]?.markdown).toBe(expected);
  } finally {
    opened.session.dispose();
  }
});
