import { expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { micromark } from 'micromark';
import { gfm, gfmHtml } from 'micromark-extension-gfm';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import { openHeadlessDocument } from '@docx-editor.dev/core/store';
import { exportMarkdown, exportMarkdownFrom, exportMarkdownLayout } from '../src/index.ts';
import { docx } from './fixture.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CODE = 'SciobotCode';

function fixture(body: string): Uint8Array {
  const entries = unzipSync(docx(body));
  entries['[Content_Types].xml'] = strToU8(
    new TextDecoder()
      .decode(entries['[Content_Types].xml'])
      .replace(
        '</Types>',
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
      )
  );
  entries['word/_rels/document.xml.rels'] = strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`
  );
  entries['word/styles.xml'] = strToU8(
    `<w:styles xmlns:w="${W}">` +
      `<w:style w:type="character" w:customStyle="1" w:styleId="${CODE}"><w:name w:val="Sciobot Code"/>` +
      '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="HTMLCode"><w:name w:val="HTML Code"/>' +
      '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:style>' +
      '</w:styles>'
  );
  return zipSync(entries);
}

const text = (value: string, props = '') =>
  `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${value}</w:t></w:r>`;
const code = (value: string, props = '') => text(value, `<w:rStyle w:val="${CODE}"/>${props}`);
const paragraph = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`;
const html = (markdown: string) =>
  micromark(markdown, {
    allowDangerousHtml: true,
    extensions: [gfm()],
    htmlExtensions: [gfmHtml()],
  });

async function project(body: string, codeCharacterStyles: readonly string[] | null = [CODE]) {
  const native = openHeadlessDocument(fixture(body));
  if (!native.ok) throw new Error(native.reason);
  const opened = openDocumentForExport(native.view);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const options = { sourceMap: true, ...(codeCharacterStyles ? { codeCharacterStyles } : {}) };
    const result = await exportMarkdownFrom(opened.session, options);
    expect(exportMarkdownLayout(await opened.session.layout(), options)).toEqual(result);
    return result;
  } finally {
    opened.session.dispose();
  }
}

test('explicit code identity projects as a code span with exact content mapping', async () => {
  const result = await project(paragraph(text('Run '), code('npm install'), text(' now.')));
  expect(result.markdown).toBe('Run `npm install` now\\.');
  expect(html(result.markdown)).toBe('<p>Run <code>npm install</code> now.</p>');
  const start = result.markdown.indexOf('npm');
  const slice = result.sourceMap!.sources.find((source) => source.markdownStart === start)!;
  expect(slice).toMatchObject({
    sourceStart: 4,
    sourceEnd: 15,
    markdownEnd: start + 11,
    exact: true,
  });
  expect(result.sourceMap!.sources.some((source) => source.markdownStart === start - 1)).toBe(
    false
  );
});

test('backtick fences and padding preserve exact code text', async () => {
  const cases: readonly [string, string, string][] = [
    ['a`b', '``a`b``', 'a`b'],
    ['`tick`', '`` `tick` ``', '`tick`'],
    ['``', '` `` `', '``'],
    ['a ``` b `', '`` a ``` b ` ``', 'a ``` b `'],
    [' padded ', '`  padded  `', ' padded '],
    ['  ', '`  `', '  '],
    ['lead ', '`lead `', 'lead '],
  ];
  for (const [value, markdown, rendered] of cases) {
    const result = await project(paragraph(text('x '), code(value), text(' y')));
    expect(result.markdown).toBe(`x ${markdown} y`);
    const expected = rendered.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    expect(html(result.markdown)).toBe(`<p>x <code>${expected}</code> y</p>`);
  }
});

test('XML characters, Markdown punctuation, and surrogate pairs stay literal inside code', async () => {
  const result = await project(paragraph(code('&lt;a href="*"&gt; &amp; _x_ \\ 😀')));
  expect(result.markdown).toBe('`<a href="*"> & _x_ \\ 😀`');
  expect(html(result.markdown)).toBe(
    '<p><code>&lt;a href=&quot;*&quot;&gt; &amp; _x_ \\ 😀</code></p>'
  );
  const slice = result.sourceMap!.sources[0]!;
  expect(slice).toMatchObject({ sourceStart: 0, sourceEnd: 23, markdownStart: 1, markdownEnd: 24 });
});

test('neighboring marks wrap code without changing its content', async () => {
  const result = await project(
    paragraph(
      text('word'),
      code('bold', '<w:b/>'),
      text(' and '),
      code('both', '<w:b/><w:i/>'),
      text(' then '),
      code('a', '<w:color w:val="FF0000"/>'),
      code('b'),
      text(' end')
    )
  );
  expect(result.markdown).toBe('word<strong>`bold`</strong> and ***`both`*** then `ab` end');
  expect(html(result.markdown)).toBe(
    '<p>word<strong><code>bold</code></strong> and <em><strong><code>both</code></strong></em> then <code>ab</code> end</p>'
  );
});

test('empty code runs, manual breaks, deletions, and links keep their intended text', async () => {
  const body =
    paragraph(text('a'), code(''), text('b')) +
    paragraph(
      code('one'),
      '<w:r><w:rPr><w:rStyle w:val="SciobotCode"/></w:rPr><w:br/><w:t>two</w:t></w:r>'
    ) +
    paragraph(
      text('keep '),
      `<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:rPr><w:rStyle w:val="${CODE}"/></w:rPr><w:delText>gone</w:delText></w:r></w:del>`
    ) +
    paragraph(
      '<w:bookmarkStart w:id="2" w:name="here"/>',
      `<w:hyperlink w:anchor="here">${code('go')}</w:hyperlink>`,
      '<w:bookmarkEnd w:id="2"/>'
    );
  const result = await project(body);
  expect(result.markdown).toBe('ab\n\n`one`  \n`two`\n\nkeep ~~`gone`~~\n\n[`go`](#here)');
  expect(html(result.markdown)).toBe(
    '<p>ab</p>\n<p><code>one</code><br />\n<code>two</code></p>\n<p>keep <del><code>gone</code></del></p>\n<p><a href="#here"><code>go</code></a></p>'
  );
});

test('GFM table cells escape pipes inside code spans', async () => {
  const cell = (content: string) => `<w:tc><w:p>${content}</w:p></w:tc>`;
  const body = `<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid><w:tr>${cell(text('H'))}${cell(text('I'))}</w:tr><w:tr>${cell(code('a|b'))}${cell(text('c'))}</w:tr></w:tbl>`;
  const result = await project(body);
  expect(result.markdown).toContain('`a\\|b`');
  expect(html(result.markdown)).toContain('<td><code>a|b</code></td>');
});

test('ordinary monospaced text and unlisted character styles stay ordinary text', async () => {
  const body = paragraph(
    text('mono', '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>'),
    text(' '),
    text('html', '<w:rStyle w:val="HTMLCode"/>'),
    text(' '),
    code('scio')
  );
  expect((await project(body)).markdown).toBe('mono html `scio`');
  expect((await project(body, null)).markdown).toBe('mono html scio');
  expect((await project(body, ['HTMLCode', CODE])).markdown).toBe('mono `html` `scio`');
  expect((await exportMarkdown(fixture(body), { codeCharacterStyles: [CODE] })).markdown).toBe(
    'mono html `scio`'
  );
});

test('malformed code style options refuse', async () => {
  for (const codeCharacterStyles of [[''], 'SciobotCode', [7]] as unknown[]) {
    await expect(
      exportMarkdown(fixture(paragraph(text('x'))), {
        codeCharacterStyles: codeCharacterStyles as readonly string[],
      })
    ).rejects.toThrow(TypeError);
  }
});

test('code-styled inline controls keep structural HTML boundaries and exact literal source text', async () => {
  const value = 'x*y`$ & <';
  const encoded = 'x*y`$ &amp; &lt;';
  const control =
    '<w:sdt><w:sdtPr><w:tag w:val="custom:literal"/></w:sdtPr><w:sdtContent>' +
    code(encoded) +
    '</w:sdtContent></w:sdt>';
  const result = await project(
    paragraph(text('Before '), code('plain'), text(' '), control, text(' after'))
  );
  expect(result.markdown).toContain('`plain`');
  expect(result.markdown).toContain('<code>');
  expect(result.markdown).toContain('</code>');
  expect(html(result.markdown)).toBe(
    '<p>Before <code>plain</code> <code>x*y`$ &amp; &lt;</code> after</p>'
  );
  const start = 'Before plain '.length;
  const source = result.sourceMap!.sources.find(
    (slice) => slice.sourceStart === start && slice.sourceEnd === start + value.length
  )!;
  expect(source).toBeDefined();
  expect(source.exact).toBe(true);
  expect(result.markdown.slice(source.markdownStart, source.markdownEnd)).not.toContain('<code>');
  expect(result.markdown.slice(source.markdownStart, source.markdownEnd)).not.toContain('</code>');
  const defaultResult = await project(paragraph(control), null);
  expect(defaultResult.markdown).not.toContain('<code>');
});
