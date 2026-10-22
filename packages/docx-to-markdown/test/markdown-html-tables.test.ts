import { expect, test } from 'bun:test';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import { exportMarkdown, exportMarkdownFrom, exportMarkdownLayout } from '../src/index.ts';
import { docx } from './fixture.ts';

const bytes = docx(
  '<w:p><w:r><w:t>Before</w:t></w:r></w:p>' +
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    '<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Header</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>A|B</w:t><w:br/><w:t>C</w:t></w:r></w:p><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>😀</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
    '<w:p><w:r><w:t>After</w:t></w:r></w:p>'
);
const html =
  'Before\n\n<table><tr><th>\n\n**Header**\n\n</th><th>\n\nB\n\n</th></tr><tr><td>\n\nA\\|B<br>C\n\nD\n\n</td><td>\n\n😀\n\n</td></tr></table>\n\nAfter';

test('HTML table mode retains header cells, marks, paragraphs, and complete source maps', async () => {
  const opened = openDocumentForExport(bytes);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const layout = await opened.session.layout();
    const before = JSON.stringify(layout);
    const result = await exportMarkdownFrom(opened.session, {
      tableSyntax: 'html',
      sourceMap: true,
    });
    expect(result.markdown).toBe(html);
    expect(exportMarkdownLayout(layout, { tableSyntax: 'html', sourceMap: true })).toEqual(result);
    const native = layout.pages
      .flatMap((page) => page.fragments)
      .find((block) => block.kind === 'table')!;
    if (native.kind !== 'table') throw new Error('No native table');
    expect(result.sourceMap!.tables[0]?.tableId).toBe(native.tableId);
    const table = result.sourceMap!.tables[0]!;
    expect(result.markdown.slice(table.markdownStart, table.markdownEnd)).toBe(html.slice(8, -7));
    expect(
      result.sourceMap!.paragraphs.map((paragraph) =>
        result.markdown.slice(paragraph.markdownStart, paragraph.markdownEnd)
      )
    ).toEqual(['Before', '**Header**', 'B', 'A\\|B<br>C', 'D', '😀', 'After']);
    const d = native.rows[1]!.cells[0]!.blocks[1]!;
    if (d.kind !== 'paragraph') throw new Error('No second cell paragraph');
    const source = result.sourceMap!.sources.find((span) => span.paragraphId === d.paragraphId)!;
    expect(result.markdown.slice(source.markdownStart, source.markdownEnd)).toBe('D');
    expect(source).toMatchObject({ sourceStart: 0, sourceEnd: 1, exact: true });
    expect(JSON.stringify(layout)).toBe(before);
    expect(() => exportMarkdownLayout(layout, { tableSyntax: 'bad' as 'html' })).toThrow(
      'tableSyntax must be gfm, html, or auto'
    );
  } finally {
    opened.session.dispose();
  }
});

test('one-shot HTML tables match session output while default and explicit GFM stay equal', async () => {
  expect((await exportMarkdown(bytes, { tableSyntax: 'html', sourceMap: true })).markdown).toBe(
    html
  );
  const plain = await exportMarkdown(bytes);
  const gfm = await exportMarkdown(bytes, { tableSyntax: 'gfm' });
  expect(gfm).toEqual(plain);
  expect(plain.markdown).toBe(
    'Before\n\n| **Header** | B |\n| --- | --- |\n| A\\|B<br>C<br>D | 😀 |\n\nAfter'
  );
});

test('auto table syntax retains merged cells and keeps independent simple tables in GFM', async () => {
  const merged = docx(
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Simple</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  );
  const opened = openDocumentForExport(merged);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const layout = await opened.session.layout(),
      before = JSON.stringify(layout);
    const result = exportMarkdownLayout(layout, { tableSyntax: 'auto', sourceMap: true });
    expect(result.markdown).toContain('<td colspan="2">\n\nMerged\n\n</td>');
    expect(result.markdown).toContain('| Simple |');
    expect(result.sourceMap!.tables).toHaveLength(2);
    const source = result.sourceMap!.sources.find(
      (value) => result.markdown.slice(value.markdownStart, value.markdownEnd) === 'Merged'
    );
    expect(source?.exact).toBe(true);
    expect(source?.sourceStart).toBe(0);
    expect(source?.sourceEnd).toBe(6);
    expect(JSON.stringify(layout)).toBe(before);
  } finally {
    opened.session.dispose();
  }
});

test('auto table syntax preserves vertical spans without duplicating continuation cells', async () => {
  const input = docx(
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Tall</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  );
  const result = await exportMarkdown(input, { tableSyntax: 'auto', sourceMap: true });
  expect(result.markdown).toContain('<td rowspan="2">\n\nTall\n\n</td>');
  expect(result.markdown.match(/<td(?: |>|$)/g)).toHaveLength(3);
  expect(result.markdown.match(/Tall/g)).toHaveLength(1);
});
