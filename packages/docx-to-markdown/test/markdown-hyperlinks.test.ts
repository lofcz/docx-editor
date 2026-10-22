import { expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { micromark } from 'micromark';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  canonicalOoxmlFingerprint,
  openHeadlessDocument,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { exportMarkdownFrom, exportMarkdownLayout } from '../src/index.ts';
import { docx } from './fixture.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
function fixture(body: string) {
  const entries = unzipSync(docx(body));
  entries['word/_rels/document.xml.rels'] = strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="safe" Type="${R}/hyperlink" Target="https://example.test/a_(b)?x=1&amp;y=2" TargetMode="External"/>` +
      `<Relationship Id="bad" Type="${R}/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>` +
      '</Relationships>'
  );
  return zipSync(entries);
}
const link = (attributes: string, content: string) =>
  `<w:hyperlink xmlns:r="${R}" ${attributes}>${content}</w:hyperlink>`;
const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
async function project(bytes: Uint8Array) {
  const native = openHeadlessDocument(bytes);
  if (!native.ok) throw new Error(native.reason);
  const pkg = native.view.currentPackage(),
    before = [...pkg.parts.values()].map(canonicalOoxmlFingerprint),
    opened = openDocumentForExport(native.view);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const result = await exportMarkdownFrom(opened.session, { sourceMap: true });
    expect(exportMarkdownLayout(await opened.session.layout(), { sourceMap: true })).toEqual(
      result
    );
    expect(native.view.currentPackage()).toBe(pkg);
    expect([...pkg.parts.values()].map(canonicalOoxmlFingerprint)).toEqual(before);
    return { result, saved: writeOoxmlPackage(pkg) };
  } finally {
    opened.session.dispose();
  }
}

test('bookmark links retain authored destinations, marks, and titles without requiring a target', async () => {
  const body =
    '<w:p><w:bookmarkStart w:id="7" w:name="present"/>' +
    link(
      'w:anchor="present" w:tooltip="A title"',
      '<w:r><w:rPr><w:b/></w:rPr><w:t>Here😀</w:t></w:r>'
    ) +
    '<w:bookmarkEnd w:id="7"/></w:p><w:p>' +
    link('w:anchor="absent"', run('Missing')) +
    '</w:p>';
  const { result, saved } = await project(fixture(body));
  expect(result.markdown).toBe('[**Here😀**](#present "A title")\n\n[Missing](#absent)');
  expect(result.pages[0]!.markdown).toBe(result.markdown);
  expect(micromark(result.markdown)).toBe(
    '<p><a href="#present" title="A title"><strong>Here😀</strong></a></p>\n' +
      '<p><a href="#absent">Missing</a></p>'
  );
  for (const text of ['Here😀', 'Missing']) {
    const start = result.markdown.indexOf(text),
      source = result.sourceMap!.sources.find((slice) => slice.markdownStart === start)!;
    expect(source).toMatchObject({
      sourceStart: 0,
      sourceEnd: text.length,
      markdownStart: start,
      markdownEnd: start + text.length,
      exact: true,
    });
  }
  expect((await project(saved)).result.markdown).toBe(result.markdown);
});

test('external titles preserve literal entity, quote, slash, and markup characters', async () => {
  const title = '&quot; &amp; &lt;tag&gt; &#92; &amp;quot; [x](#fake)',
    { result, saved } = await project(
      fixture('<w:p>' + link(`r:id="safe" w:tooltip="${title}"`, run('Visit')) + '</w:p>')
    );
  expect(result.markdown).toBe(
    '[Visit](https://example.test/a_%28b%29?x=1&y=2 "&quot; &amp; &lt;tag&gt; &#92; &amp;quot; [x](#fake)")'
  );
  expect(micromark(result.markdown)).toBe(
    '<p><a href="https://example.test/a_%28b%29?x=1&amp;y=2" ' +
      'title="&quot; &amp; &lt;tag&gt; \\ &amp;quot; [x](#fake)">Visit</a></p>'
  );
  expect((await project(saved)).result.markdown).toBe(result.markdown);
});

test('empty titles stay distinct from absent titles and unsafe or missing relationships stay inert', async () => {
  const { result } = await project(
    fixture(
      '<w:p>' +
        link('w:anchor="here" w:tooltip=""', run('Empty')) +
        link('w:anchor="here"', run('Absent')) +
        link('r:id="bad" w:tooltip="unsafe"', run('Bad')) +
        link('r:id="missing" w:tooltip="missing"', run('Missing')) +
        '</w:p>'
    )
  );
  expect(result.markdown).toBe('[Empty](#here "")[Absent](#here)BadMissing');
  expect(micromark(result.markdown)).toBe(
    '<p><a href="#here">Empty</a><a href="#here">Absent</a>BadMissing</p>'
  );
  expect(result.markdown).not.toContain('javascript:');
});
