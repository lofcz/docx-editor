import { expect, test } from 'bun:test';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  openHeadlessDocument,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { exportMarkdown, exportMarkdownFrom, exportMarkdownLayout } from '../src/index.ts';
import { docx } from './fixture.ts';
const ID = 'OwnedRule';
const rule = (style = ID, content = '', edges = '') =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1"/><w:between w:val="single" w:sz="6" w:space="1"/>${edges}</w:pBdr></w:pPr>${content}</w:p>`;
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const options = { sourceMap: true, horizontalRuleParagraphStyles: [ID] };

test('opted-in empty native bottom border exports one rule with an exact generated paragraph extent through every entry point', async () => {
  const bytes = docx(p('Before') + rule() + p('After'));
  const native = openHeadlessDocument(bytes);
  if (!native.ok) throw new Error(native.reason);
  const opened = openDocumentForExport(native.view);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const result = await exportMarkdownFrom(opened.session, options);
    expect(result.markdown).toBe('Before\n\n---\n\nAfter');
    expect(exportMarkdownLayout(await opened.session.layout(), options)).toEqual(result);
    expect((await exportMarkdown(bytes, options)).markdown).toBe(result.markdown);
    expect((await exportMarkdown(bytes)).markdown).toBe('Before\n\n\n\nAfter');
    const extent = result.sourceMap!.paragraphs[1]!;
    expect(extent).toMatchObject({
      sourceStart: 0,
      sourceEnd: 0,
      markdownStart: 8,
      markdownEnd: 11,
    });
    expect(
      result.sourceMap!.sources.filter((source) => source.paragraphId === extent.paragraphId)
    ).toHaveLength(0);
    const read = readOoxmlPackage(bytes);
    if (!read.ok) throw new Error(read.reason);
    expect((await exportMarkdown(writeOoxmlPackage(read.package), options)).markdown).toBe(
      result.markdown
    );
  } finally {
    opened.session.dispose();
  }
});

test('ordinary imported borders and nonempty or boxed style collisions never become thematic breaks', async () => {
  for (const body of [
    rule('Imported'),
    rule(ID, '<w:r><w:t>Keep</w:t></w:r>'),
    rule(ID, '', '<w:top w:val="single" w:sz="6"/>'),
    rule().replaceAll('w:val="single"', 'w:val="none"'),
  ]) {
    const result = await exportMarkdown(docx(body), options);
    expect(result.markdown).not.toContain('---');
    if (body.includes('Keep')) expect(result.markdown).toBe('Keep');
  }
});

test('rule option rejects invalid style lists before translation', async () => {
  const native = openHeadlessDocument(docx(rule()));
  if (!native.ok) throw new Error(native.reason);
  const opened = openDocumentForExport(native.view);
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const layout = await opened.session.layout();
    for (const styles of ['', [''], [5]] as unknown[])
      expect(() =>
        exportMarkdownLayout(layout, { horizontalRuleParagraphStyles: styles as readonly string[] })
      ).toThrow('horizontalRuleParagraphStyles');
  } finally {
    opened.session.dispose();
  }
});

test('consecutive rule paragraphs retain distinct zero-width owners', async () => {
  const result = await exportMarkdown(docx(rule() + rule()), options);
  expect(result.markdown).toBe('---\n\n---');
  expect(result.sourceMap!.paragraphs).toHaveLength(2);
  expect(
    result.sourceMap!.paragraphs.map((p) => [
      p.sourceStart,
      p.sourceEnd,
      p.markdownStart,
      p.markdownEnd,
    ])
  ).toEqual([
    [0, 0, 0, 3],
    [0, 0, 5, 8],
  ]);
});

test('native paint carries exactly one closing stroke per single, adjacent, and page-separated marker', async () => {
  for (const body of [
    rule(),
    rule() + rule(),
    rule() + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + rule(),
  ]) {
    const opened = openDocumentForExport(docx(body));
    if (!opened.ok) throw new Error(opened.reason);
    try {
      const layout = await opened.session.layout();
      const paragraphs = layout.pages
        .flatMap((page) => page.fragments)
        .filter((record) => record.kind === 'paragraph' && record.styleId === ID);
      const strokes = paragraphs.flatMap((paragraph) => paragraph.borders ?? []);
      expect(strokes).toHaveLength(body === rule() ? 1 : 2);
      expect(
        strokes.every(
          (stroke) =>
            ['bottom', 'between'].includes(stroke.side) &&
            stroke.edge.widthPt === 0.75 &&
            stroke.box.width > 0
        )
      ).toBe(true);
      expect(new Set(paragraphs.map((paragraph) => paragraph.paragraphId)).size).toBe(
        body === rule() ? 1 : 2
      );
      expect(exportMarkdownLayout(layout, options).markdown.match(/---/g)).toHaveLength(
        body === rule() ? 1 : 2
      );
      if (body.includes('w:type="page"')) expect(layout.pages.length).toBeGreaterThan(1);
    } finally {
      opened.session.dispose();
    }
  }
});

test('a rule split into two page fragments retains one closing stroke and one logical owner', async () => {
  const opened = openDocumentForExport(docx(rule()));
  if (!opened.ok) throw new Error(opened.reason);
  try {
    const original = await opened.session.layout();
    const page = original.pages[0]!;
    const paragraph = page.fragments.find((fragment) => fragment.kind === 'paragraph');
    if (!paragraph || paragraph.kind !== 'paragraph') throw new Error('Missing rule paragraph');
    // Empty authored rules are one line. Exercise the detached layout continuation contract
    // independently: only the final continuation carries the native closing stroke.
    const first = { ...paragraph, paragraphEnd: undefined, bottomBorder: undefined, borders: [] };
    const last = { ...paragraph, id: paragraph.id + '-continued', fragmentIndex: 1 };
    const split = {
      ...original,
      pages: [
        { ...page, fragments: [first] },
        { ...page, id: page.id + '-continued', index: 1, fragments: [last] },
      ],
    };
    const result = exportMarkdownLayout(split, options);
    expect(
      split.pages.flatMap((item) => item.fragments).flatMap((item) => item.borders ?? [])
    ).toHaveLength(1);
    expect(result.markdown).toBe('---');
    expect(result.sourceMap!.paragraphs).toHaveLength(1);
    expect(result.sourceMap!.paragraphs[0]).toMatchObject({
      sourceStart: 0,
      sourceEnd: 0,
      markdownStart: 0,
      markdownEnd: 3,
    });
    expect(result.sourceMap!.sources).toEqual([]);
  } finally {
    opened.session.dispose();
  }
});
