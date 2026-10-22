import { expect, test } from 'bun:test';
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkMdx from 'remark-mdx';
import { openDocumentForExport, type ExportSemanticLayout } from '@docx-editor.dev/core/export';
import { escapeXml } from '@docx-editor.dev/core/store';
import reference from '../__tests__/fixtures/ai-pairs.json';
import { docxFixture, W } from '../../test/support/document';
import { findAndReplace } from '../tools/find-replace';
import { markdownPointAt, markdownRangeAt, projectMarkdown } from './markdown';

const revision = { documentId: 'lesson-1', epoch: 'mount-1', revision: 3 };
async function layoutOf(bytes: Uint8Array): Promise<ExportSemanticLayout> {
  const opened = openDocumentForExport(bytes);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    return await opened.session.layout();
  } finally {
    opened.session.dispose();
  }
}
const paragraph = (text: string) => `<w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;

interface FormattingBlock {
  readonly type: string;
  readonly listStyleType?: string;
  readonly children: readonly { text: string; bold?: boolean; italic?: boolean }[];
}
// Only this captured formatting case uses this test builder. Production migration stays in I11.
function formattingDocx(blocks: readonly FormattingBlock[]) {
  const numbering =
    `<w:numbering xmlns:w="${W}">` +
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>';
  const body = blocks
    .map(
      (block) =>
        '<w:p>' +
        (block.listStyleType
          ? `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${block.listStyleType === 'decimal' ? '1' : '2'}"/></w:numPr></w:pPr>`
          : '') +
        block.children
          .map(
            (child) =>
              `<w:r><w:rPr>${child.bold ? '<w:b/>' : ''}${child.italic ? '<w:i/>' : ''}</w:rPr><w:t xml:space="preserve">${escapeXml(child.text)}</w:t></w:r>`
          )
          .join('') +
        '</w:p>'
    )
    .join('');
  return docxFixture(body, { numbering });
}

test('C01 formatting and lists produce stable, copyable native Markdown', async () => {
  const captured = reference.cases.find((value) => value.name === 'formatting and lists')!;
  for (const [blocks, expected] of [
    [
      captured.canonicalBefore,
      '**Bold** and *italic*\\.\n\n1. First\n\n2. Second\n\n- Apple\n\n- Pear',
    ],
    [
      captured.expected,
      '**Bold** and *updated italic*\\.\n\n1. First\n\n2. Added\n\n3. Second\n\n- Apple\n\n- Pear',
    ],
  ] as const) {
    const layout = await layoutOf(formattingDocx(blocks as FormattingBlock[]));
    const before = JSON.stringify(layout);
    const projection = projectMarkdown(layout, revision);
    expect(projection.text).toBe(expected);
    expect(
      projection.sourceMap.paragraphs.map((paragraph) =>
        projection.text.slice(paragraph.markdownStart, paragraph.markdownEnd)
      )
    ).toEqual(expected.split('\n\n'));
    expect(findAndReplace(projection.text, projection.text, 'Copied', true)).toEqual({
      success: true,
      newContent: 'Copied',
    });
    expect(JSON.stringify(layout)).toBe(before);
    expect(projection.revision).toEqual(revision);
    expect(projection.revision === revision).toBe(false);
    expect(JSON.parse(JSON.stringify(projection))).toEqual(projection);
  }
});

test('marks and heading prefixes have no native text offsets', async () => {
  const layout = await layoutOf(
    docxFixture(
      '<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p>'
    )
  );
  const projection = projectMarkdown(layout, revision);
  expect(projection.text).toBe('## Title\n\n**Bold**');
  for (const position of [0, 1, 2, 9, 10, 11, 17])
    expect(markdownPointAt(projection, position)).toBeNull();
  expect(markdownRangeAt(projection, 12, 16)).toMatchObject({
    start: { offset: 0 },
    end: { offset: 4 },
  });
  expect(projection.syntax.map(({ start, end }) => projection.text.slice(start, end))).toEqual([
    '## ',
    '\n\n**',
    '**',
  ]);
});

test('escape and emoji interiors never resolve to editable native positions', async () => {
  const projection = projectMarkdown(await layoutOf(docxFixture(paragraph('A😀&*B'))), revision);
  expect(projection.text).toBe('A😀&amp;\\*B');
  for (const position of [2, 4, 5, 6, 7, 9, -1, 1.5, Infinity, 12])
    expect(markdownPointAt(projection, position)).toBeNull();
  expect(markdownPointAt(projection, 3)).toMatchObject({ sourceScope: 'body', offset: 3 });
  expect(markdownPointAt(projection, 8)).toMatchObject({ offset: 4 });
  expect(markdownPointAt(projection, 10)).toMatchObject({ offset: 5 });
  expect(markdownRangeAt(projection, 3, 10)).toMatchObject({
    start: { offset: 3 },
    end: { offset: 5 },
  });
  expect(markdownRangeAt(projection, 10, 3)).toBeNull();
});

test('quote prefixes and link destinations remain generated syntax', async () => {
  const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const projection = projectMarkdown(
    await layoutOf(
      docxFixture(
        '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>Same</w:t></w:r></w:p>' +
          '<w:p><w:hyperlink r:id="link"><w:r><w:t>Same</w:t></w:r></w:hyperlink></w:p>',
        {
          relationships: `<Relationship Id="link" Type="${relationship}/hyperlink" Target="https://example.com" TargetMode="External"/>`,
        }
      )
    ),
    revision
  );
  expect(projection.text).toBe('> Same\n\n[Same](https://example.com)');
  expect(markdownPointAt(projection, 0)).toBeNull();
  expect(markdownPointAt(projection, 1)).toBeNull();
  expect(markdownPointAt(projection, 18)).toBeNull();
  const quote = markdownPointAt(projection, 2)!;
  const link = markdownPointAt(projection, 9)!;
  expect(quote.offset).toBe(0);
  expect(link.offset).toBe(0);
  expect(quote.paragraphId === link.paragraphId).toBe(false);
});

test('the selected native review layout retains original source provenance', async () => {
  const bytes = docxFixture(
    '<w:p><w:r><w:t>A</w:t></w:r><w:del w:id="1" w:author="Ada"><w:r><w:delText>Old</w:delText></w:r></w:del>' +
      '<w:ins w:id="2" w:author="Ada"><w:r><w:t>New</w:t></w:r></w:ins><w:r><w:t>Z</w:t></w:r></w:p>'
  );
  const opened = openDocumentForExport(bytes);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const original = projectMarkdown(await opened.session.layoutFor('original'), revision);
    const proposed = projectMarkdown(await opened.session.layoutFor('proposed'), revision);
    expect(original.text).toBe('AOldZ');
    expect(proposed.text).toBe('ANewZ');
    expect(original.displayMode).toBe('original');
    expect(proposed.displayMode).toBe('proposed');
    expect(markdownPointAt(original, 4)).toBeNull();
    expect(markdownPointAt(proposed, 1)).toBeNull();
    expect(markdownPointAt(original, 4, 'backward')).toMatchObject({ offset: 4 });
    expect(markdownPointAt(original, 4, 'forward')).toMatchObject({ offset: 7 });
    expect(markdownRangeAt(original, 1, 4)).toMatchObject({
      start: { offset: 1 },
      end: { offset: 4 },
    });
    expect(markdownRangeAt(proposed, 1, 4)).toMatchObject({
      start: { offset: 4 },
      end: { offset: 7 },
    });
  } finally {
    opened.session.dispose();
  }
});

test('HTML table projection accepts multiline MDX cell content without losing paragraphs or marks', async () => {
  const layout = await layoutOf(
    docxFixture(
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>WIDGET</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>After</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )
  );
  const projection = projectMarkdown(layout, revision, { tableSyntax: 'html' });
  const text = projection.text.replace(
    'WIDGET',
    '<listening title="Lesson">\n\n<line speaker="s1">Hello</line>\n\n</listening>'
  );
  const tree = unified().use(remarkParse).use(remarkMdx).parse(text);
  const names: string[] = [];
  const types: string[] = [];
  const visit = (node: { type: string; name?: string | null; children?: unknown[] }) => {
    if (node.name) names.push(node.name);
    types.push(node.type);
    for (const child of node.children ?? []) visit(child as typeof node);
  };
  visit(tree);
  expect(names).toEqual(['table', 'tr', 'td', 'listening', 'line']);
  expect(types).toContain('strong');
  expect(types.filter((type) => type === 'paragraph')).toHaveLength(2);
  const after = markdownPointAt(projection, projection.text.indexOf('After'));
  expect(after?.offset).toBe(0);
});
