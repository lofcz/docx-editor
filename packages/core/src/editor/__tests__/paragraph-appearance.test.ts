import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createDocxEditor } from '../docx-editor.ts';
import { openDocumentForExport } from '../../export/export-session.ts';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import {
  createFixedMeasurer,
  paragraphFragmentsOf,
  paragraphFragmentsOfBlocks,
} from '@docx-editor.dev/core/layout';
import type { ParagraphAppearanceContribution } from '../../contracts/paragraph-appearance.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const paragraph = (label: string) =>
  `<w:p><w:sdt><w:sdtPr><w:tag w:val="test:task"/><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>${label}</w:t></w:r></w:p>`;
const contribution: ParagraphAppearanceContribution = {
  id: 'task-test',
  resolve: ({ controls }) =>
    controls.some((control) => control.tag === 'test:task' && control.checked)
      ? { strike: true, color: { val: '606060' } }
      : null,
};
function documentBytes() {
  const namespaces = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Type="${R}/header" Target="header.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document ${namespaces}><w:body>${paragraph('Body')}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${paragraph('Cell')}</w:tc></w:tr></w:tbl><w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr></w:body></w:document>`
    ),
    'word/header.xml': strToU8(`<w:hdr ${namespaces}>${paragraph('Header')}</w:hdr>`),
  });
}
test('native paint, effective formatting, header/table layout, export, and raw reopen share appearance without model writes', async () => {
  const bytes = documentBytes(),
    container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: bytes,
    modules: [{ id: 'tasks', paragraphAppearance: contribution }],
  });
  try {
    for (const label of ['Body', 'Cell', 'Header']) {
      const span = [...container.querySelectorAll<HTMLElement>('span')].find(
        (node) => node.textContent === label
      );
      expect(span?.style.textDecorationLine).toBe('line-through');
      expect(['#606060', 'rgb(96, 96, 96)']).toContain(span?.style.color);
    }
    const surface = editor.surface!,
      id = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 2 },
      head: { paragraphId: id, offset: 2 },
    });
    expect(surface.formatting().strikethrough).toBe(true);
    const opened = openDocumentForExport(surface.session, {
      paragraphAppearance: contribution,
      measurer: createFixedMeasurer(),
      displayMode: 'proposed',
    });
    if (!opened.ok) throw new Error(opened.reason);
    try {
      const layout = await opened.session.layout();
      const spans = layout.pages
        .flatMap((page) => [
          ...paragraphFragmentsOf(page),
          ...paragraphFragmentsOfBlocks(page.header?.fragments ?? []),
        ])
        .flatMap((p) => p.lines.flatMap((line) => line.spans));
      for (const label of ['Body', 'Cell', 'Header'])
        expect(spans.find((span) => span.text === label)?.style).toMatchObject({
          strike: true,
          color: '606060',
        });
    } finally {
      opened.session.dispose();
    }
    const before = surface.session.currentPackage();
    const saved = new Uint8Array(await editor.save()),
      after = readOoxmlPackage(saved);
    if (!after.ok) throw new Error('Cannot read fixture');
    for (const name of ['/word/document.xml', '/word/header.xml'])
      expect(canonicalOoxmlFingerprint(after.package.parts.get(name)!.root)).toBe(
        canonicalOoxmlFingerprint(before.parts.get(name)!.root)
      );
    await editor.load(saved);
    expect(
      [...container.querySelectorAll<HTMLElement>('span')].find(
        (node) => node.textContent === 'Body'
      )?.style.textDecorationLine
    ).toBe('line-through');
  } finally {
    editor.destroy();
    container.remove();
  }
});
