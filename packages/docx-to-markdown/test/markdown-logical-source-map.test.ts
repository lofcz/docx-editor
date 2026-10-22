import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  exportMarkdown,
  exportMarkdownFrom,
  exportMarkdownLayout,
  toMarkdownJSON,
  type MarkdownSourceMap,
} from '../src/index.ts';
import { docx } from './fixture.ts';

function markdownOffset(map: MarkdownSourceMap, paragraphId: string, offset: number): number[] {
  const found = new Set<number>();
  for (const slice of map.sources) {
    if (!slice.exact || slice.paragraphId !== paragraphId) continue;
    const boundaries = slice.markdownBoundaries ?? [];
    for (const [index, boundary] of boundaries.entries()) {
      if (boundary.sourceOffset === offset) found.add(boundary.markdownOffset);
      const next = boundaries[index + 1];
      if (
        next &&
        boundary.sourceOffset < offset &&
        offset < next.sourceOffset &&
        next.sourceOffset - boundary.sourceOffset === next.markdownOffset - boundary.markdownOffset
      ) {
        found.add(boundary.markdownOffset + offset - boundary.sourceOffset);
      }
    }
  }
  return [...found];
}

async function project(body: string | Uint8Array) {
  const opened = openDocumentForExport(typeof body === 'string' ? docx(body) : body);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const layout = await opened.session.layout();
    const before = JSON.stringify(layout);
    const result = await exportMarkdownFrom(opened.session, { sourceMap: true });
    const detached = exportMarkdownLayout(layout, { sourceMap: true });
    expect(detached).toEqual(result);
    expect(JSON.stringify(layout)).toBe(before);
    return result;
  } finally {
    opened.session.dispose();
  }
}

test('maps every ordinary UTF-16 boundary through escaping without review artifacts', async () => {
  const result = await project(
    '<w:p><w:r><w:t>č😀* &amp;&lt;&gt;</w:t><w:tab/><w:t>Z</w:t></w:r></w:p>'
  );
  expect(result.markdown).toBe('č😀\\* &amp;&lt;&gt; Z');
  expect(result.reviewArtifacts).toEqual([]);
  const map = result.sourceMap!;
  expect(map.unit).toBe('utf16-code-unit');
  const paragraph = map.paragraphs[0]!;
  expect(paragraph).toMatchObject({ sourceScope: 'body', sourceStart: 0, sourceEnd: 10 });
  const expected = [0, 1, 2, 3, 5, 6, 11, 15, 19, 20, 21];
  for (const [source, markdown] of expected.entries()) {
    expect(markdownOffset(map, paragraph.paragraphId, source)).toEqual([markdown]);
  }
  // Escape interiors cannot be mistaken for independently editable source characters.
  expect(
    map.sources.some((slice) =>
      slice.markdownBoundaries?.some((boundary) => boundary.markdownOffset === 7)
    )
  ).toBe(false);
  expect(Object.isFrozen(map)).toBe(true);
  expect(Object.isFrozen(map.sources)).toBe(true);
  expect(Object.isFrozen(map.sources[0]?.markdownBoundaries?.[0])).toBe(true);
  expect(JSON.parse(JSON.stringify(toMarkdownJSON(result))).sourceMap).toEqual(map);
});

test('keeps generated heading and mark delimiters outside native text spans', async () => {
  const result = await project(
    '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>' +
      '<w:r><w:t xml:space="preserve"> and </w:t></w:r><w:r><w:rPr><w:i/><w:strike/></w:rPr><w:t>Other</w:t></w:r></w:p>'
  );
  expect(result.markdown).toBe('# Title\n\n**Bold** and *~~Other~~*');
  const map = result.sourceMap!;
  expect(
    map.sources.map((slice) => result.markdown.slice(slice.markdownStart, slice.markdownEnd))
  ).toEqual(['Title', 'Bold', ' and ', 'Other']);
  expect(map.sources.some((slice) => slice.markdownStart < 2)).toBe(false);
  expect(markdownOffset(map, map.paragraphs[1]!.paragraphId, 2)).toEqual([13]);
  expect(
    map.paragraphs.map((paragraph) =>
      result.markdown.slice(paragraph.markdownStart, paragraph.markdownEnd)
    )
  ).toEqual(['# Title', '**Bold** and *~~Other~~*']);
});

test('maps hyperlink labels without mapping generated destinations', async () => {
  const relationshipNamespace =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const entries = unzipSync(
    docx(
      `<w:p><w:hyperlink xmlns:r="${relationshipNamespace}" r:id="web"><w:r><w:t>Go</w:t></w:r></w:hyperlink></w:p>` +
        '<w:p><w:r><w:t>Target</w:t></w:r></w:p>'
    )
  );
  entries['word/_rels/document.xml.rels'] = strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="web" Type="${relationshipNamespace}/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>`
  );
  const result = await project(zipSync(entries));
  expect(result.markdown).toBe('[Go](https://example.com)\n\nTarget');
  const map = result.sourceMap!;
  expect(
    map.sources.map((slice) => result.markdown.slice(slice.markdownStart, slice.markdownEnd))
  ).toEqual(['Go', 'Target']);
  expect(markdownOffset(map, map.paragraphs[0]!.paragraphId, 0)).toEqual([1]);
  expect(markdownOffset(map, map.paragraphs[0]!.paragraphId, 2)).toEqual([3]);
  expect(map.sources.some((slice) => slice.markdownStart <= 7 && 7 < slice.markdownEnd)).toBe(
    false
  );
});

test('preserves Word quote styles, marks, and hard breaks with exact text boundaries', async () => {
  const result = await project(
    '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>First</w:t></w:r>' +
      '<w:r><w:br/><w:t>Second</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="IntenseQuote"/></w:pPr><w:r><w:t>Third</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="CustomQuote"/></w:pPr><w:r><w:t>Plain</w:t></w:r></w:p>'
  );
  expect(result.markdown).toBe('> **First**  \n> Second\n\n> Third\n\n>\n\nPlain');
  const map = result.sourceMap!;
  expect(map.paragraphs).toHaveLength(4);
  expect(markdownOffset(map, map.paragraphs[0]!.paragraphId, 0)).toEqual([4]);
  expect(markdownOffset(map, map.paragraphs[0]!.paragraphId, 7)).toEqual([17]);
  expect(markdownOffset(map, map.paragraphs[1]!.paragraphId, 0)).toEqual([26]);
  expect(
    map.paragraphs.map((paragraph) =>
      result.markdown.slice(paragraph.markdownStart, paragraph.markdownEnd)
    )
  ).toEqual(['> **First**  \n> Second', '> Third', '>', 'Plain']);
  expect(map.sources.some((slice) => slice.markdownStart < 2)).toBe(false);
});

test('joins consecutive paragraphs of one quote style into one blockquote with exact maps', async () => {
  const quote = (style: string, text: string) =>
    `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const result = await project(
    quote('Quote', 'One') +
      quote('Quote', 'Two') +
      quote('IntenseQuote', 'Loud') +
      '<w:p><w:r><w:t>Between</w:t></w:r></w:p>' +
      quote('Quote', 'Three') +
      '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr></w:p>' +
      quote('Quote', 'Four')
  );
  expect(result.markdown).toBe('> One\n>\n> Two\n\n> Loud\n\nBetween\n\n> Three\n>\n>\n>\n> Four');
  const map = result.sourceMap!;
  expect(
    map.paragraphs.map((paragraph) =>
      result.markdown.slice(paragraph.markdownStart, paragraph.markdownEnd)
    )
  ).toEqual(['> One', '> Two', '> Loud', 'Between', '> Three', '>', '> Four']);
  const ids = map.paragraphs.map((paragraph) => paragraph.paragraphId);
  expect(markdownOffset(map, ids[0]!, 0)).toEqual([2]);
  expect(markdownOffset(map, ids[1]!, 0)).toEqual([10]);
  expect(markdownOffset(map, ids[1]!, 3)).toEqual([13]);
  expect(markdownOffset(map, ids[2]!, 0)).toEqual([17]);
  expect(markdownOffset(map, ids[4]!, 0)).toEqual([34]);
  expect(markdownOffset(map, ids[6]!, 4)).toEqual([52]);
  // Generated separator lines never map to a native paragraph.
  for (const offset of [5, 6, 7])
    expect(
      map.sources.some((slice) => slice.markdownStart <= offset && offset < slice.markdownEnd)
    ).toBe(false);
});

test('maps table cell text, escaped pipes, and native hard breaks', async () => {
  const result = await project(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>' +
      '<w:tc><w:p><w:r><w:t>A|B</w:t><w:br/><w:t>😀</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>A|B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  );
  expect(result.markdown).toBe('| A\\|B<br>😀 | A\\|B |\n| --- | --- |');
  const map = result.sourceMap!;
  expect(map.paragraphs).toHaveLength(2);
  const first = map.paragraphs[0]!.paragraphId;
  const second = map.paragraphs[1]!.paragraphId;
  expect(first === second).toBe(false);
  for (const [source, markdown] of [2, 3, 5, 6, 10, 11, 12].entries()) {
    expect(markdownOffset(map, first, source)).toEqual([markdown]);
  }
  expect(markdownOffset(map, second, 0)).toEqual([15]);
  expect(
    map.paragraphs.map((paragraph) =>
      result.markdown.slice(paragraph.markdownStart, paragraph.markdownEnd)
    )
  ).toEqual(['A\\|B<br>😀', 'A\\|B']);
});

test('keeps plain text maps compact and retains empty source paragraphs', async () => {
  const result = await project(
    '<w:p/>' + '<w:p><w:r><w:t>' + 'abc '.repeat(200) + '</w:t></w:r></w:p><w:p/>'
  );
  const map = result.sourceMap!;
  expect(map.paragraphs).toHaveLength(3);
  expect(map.paragraphs.map(({ sourceStart, sourceEnd }) => [sourceStart, sourceEnd])).toEqual([
    [0, 0],
    [0, 800],
    [0, 0],
  ]);
  expect(map.sources.flatMap((slice) => slice.markdownBoundaries ?? [])).toHaveLength(2);
  expect(markdownOffset(map, map.paragraphs[1]!.paragraphId, 411)).toEqual([413]);
  expect(
    map.paragraphs.map(({ markdownStart, markdownEnd }) => [markdownStart, markdownEnd])
  ).toEqual([
    [0, 0],
    [2, 802],
    [802, 802],
  ]);
  const blank = await project('<w:p/>');
  expect(blank.markdown).toBe('');
  expect(blank.sourceMap?.paragraphs).toHaveLength(1);
  expect(blank.sourceMap?.paragraphs[0]).toMatchObject({ markdownStart: 0, markdownEnd: 0 });
});

test('one-shot opt-in preserves default output and omits source maps by default', async () => {
  const bytes = docx('<w:p><w:r><w:t>Keep *this*</w:t></w:r></w:p>');
  const ordinary = await exportMarkdown(bytes);
  const mapped = await exportMarkdown(bytes, { sourceMap: true });
  expect(Object.hasOwn(ordinary, 'sourceMap')).toBe(false);
  const { sourceMap, ...unchanged } = mapped;
  expect(unchanged).toEqual(ordinary);
  expect(sourceMap?.sources).toHaveLength(1);
});

test('paragraph constructs stay exact through nested table HTML wrappers', async () => {
  const result = await project(
    '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
      '<w:p><w:r><w:t>Before</w:t></w:r></w:p>' +
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      '<w:p><w:r><w:t>After</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  );
  expect(result.markdown).toContain('<table>');
  expect(
    result.sourceMap!.paragraphs.map((paragraph) =>
      result.markdown.slice(paragraph.markdownStart, paragraph.markdownEnd)
    )
  ).toEqual(['Before', '**Inner**', 'After']);
  expect(result.sourceMap!.tables).toHaveLength(2);
  const [outer, inner] = result.sourceMap!.tables;
  expect(result.markdown.slice(outer!.markdownStart, outer!.markdownEnd)).toBe(result.markdown);
  expect(result.markdown.slice(inner!.markdownStart, inner!.markdownEnd)).toBe(
    '<table><tr><td>**Inner**</td></tr></table>'
  );
});

test('paragraph constructs stay exact through note definition prefixes and indentation', async () => {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const entries = unzipSync(
    docx('<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="1"/></w:r></w:p>')
  );
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'
    )
  );
  entries['word/_rels/document.xml.rels'] = strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="notes" Type="${r}/footnotes" Target="footnotes.xml"/></Relationships>`
  );
  entries['word/footnotes.xml'] = strToU8(
    `<w:footnotes xmlns:w="${w}"><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>First</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>N</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:footnote></w:footnotes>`
  );
  const result = await project(zipSync(entries));
  expect(result.markdown).toBe(
    'Body[^1]\n\n[^1]: First\n    \n    | N |\n    | --- |\n    \n    Second'
  );
  expect(
    result.sourceMap!.paragraphs.map((paragraph) =>
      result.markdown.slice(paragraph.markdownStart, paragraph.markdownEnd)
    )
  ).toEqual(['Body[^1]', 'First', 'N', 'Second']);
  expect(result.sourceMap!.paragraphs[1]?.sourceScope.startsWith('footnote:')).toBe(true);
  const table = result.sourceMap!.tables[0]!;
  expect(table.sourceScope.startsWith('footnote:')).toBe(true);
  expect(result.markdown.slice(table.markdownStart, table.markdownEnd)).toBe('| N |\n    | --- |');
});

test('complete empty-cell tables retain canonical IDs and all generated separators', async () => {
  const bytes = docx(
    '<w:p><w:r><w:t>Before</w:t></w:r></w:p>' +
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:p><w:r><w:t>After</w:t></w:r></w:p>'
  );
  const opened = openDocumentForExport(bytes);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const layout = await opened.session.layout();
    const tableRecord = layout.pages
      .flatMap((page) => page.fragments)
      .find((block) => block.kind === 'table')!;
    if (tableRecord.kind !== 'table') throw new Error('No native table');
    const result = await exportMarkdownFrom(opened.session, { sourceMap: true });
    const table = result.sourceMap!.tables[0]!;
    expect(result.sourceMap!.tables).toHaveLength(1);
    expect(table.tableId).toBe(tableRecord.tableId);
    expect(table.sourceScope).toBe('body');
    expect(result.markdown.slice(table.markdownStart, table.markdownEnd)).toBe(
      '|  |  |\n| --- | --- |'
    );
    expect(result.markdown).toBe('Before\n\n|  |  |\n| --- | --- |\n\nAfter');
    expect(Object.isFrozen(table)).toBe(true);
    expect(Object.isFrozen(result.sourceMap!.tables)).toBe(true);
    expect((await exportMarkdownFrom(opened.session)).markdown).toBe(result.markdown);
  } finally {
    opened.session.dispose();
  }
});
