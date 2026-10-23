import { expect, test } from 'bun:test';
import { createFixedMeasurer } from '@docx-editor.dev/core/layout';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { exportMarkdownFrom, exportMarkdownLayout, openDocumentForExport } from '../src/index.ts';
import { imageDocx, picture, PNG, R } from './media-fixture.ts';

test('detached images retain native spans beside hyperlinks after session disposal', async () => {
  const body = `<w:p><w:hyperlink xmlns:r="${R}" r:id="rLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink><w:r>${picture()}<w:t>Tail</w:t></w:r></w:p>`;
  const entries = unzipSync(imageDocx(body));
  entries['word/_rels/document.xml.rels'] = strToU8(
    strFromU8(entries['word/_rels/document.xml.rels']!).replace(
      '</Relationships>',
      `<Relationship Id="rLink" Type="${R}/hyperlink" Target="https://example.test" TargetMode="External"/></Relationships>`
    )
  );
  const opened = await openDocumentForExport(zipSync(entries), {
    measurer: createFixedMeasurer(),
  });
  if (!opened.ok) throw Error(opened.reason);
  const { session } = opened;
  let layout;
  let live;
  try {
    layout = await session.layout();
    live = await exportMarkdownFrom(session, { images: true, sourceMap: true });
  } finally {
    session.dispose();
  }
  const result = exportMarkdownLayout(layout, { imageAssets: live.media, sourceMap: true });
  expect(result.markdown).toBe(`[Link](https://example.test)![Picture](${live.media[0]!.url})Tail`);
  expect(result.markdown).toBe(live.markdown);
  expect(result.sourceMap).toEqual(live.sourceMap);
  expect(result.media).toEqual(live.media);
  expect(result.media[0]!.bytes).toEqual(PNG);
  const occurrence = live.media[0]!.occurrences[0]!;
  const start = result.markdown.indexOf('![Picture]');
  const end = result.markdown.indexOf('Tail');
  const span = result.sourceMap!.sources.find(
    (source) =>
      source.paragraphId === occurrence.paragraphId && source.sourceStart === occurrence.start
  );
  expect(span).toMatchObject({
    sourceEnd: occurrence.start + 1,
    markdownStart: start,
    markdownEnd: end,
  });
  expect(result.warnings).toEqual([]);
  const omitted = exportMarkdownLayout(layout);
  expect(omitted.media).toEqual([]);
  expect(omitted.markdown).not.toContain('![Picture]');
  expect(omitted.warnings.some((warning) => warning.code === 'omitted-drawing')).toBe(true);
});

test('a supplied occurrence subset leaves other drawings omitted', async () => {
  const opened = await openDocumentForExport(
    imageDocx(
      `<w:p><w:r>${picture(1, false, 'First')}<w:t>Between</w:t>${picture(2, false, 'Second')}</w:r></w:p>`
    ),
    { measurer: createFixedMeasurer() }
  );
  if (!opened.ok) throw Error(opened.reason);
  const { session } = opened;
  try {
    const layout = await session.layout();
    const live = await exportMarkdownFrom(session, { images: true });
    expect(live.media[0]!.occurrences).toHaveLength(2);
    const asset = live.media[0]!;
    const result = exportMarkdownLayout(layout, {
      imageAssets: [
        {
          ...asset,
          url: 'https://cdn.example.test/first.png',
          occurrences: [asset.occurrences[0]!],
        },
      ],
    });
    expect(result.markdown).toContain('![First](https://cdn.example.test/first.png)');
    expect(result.markdown).not.toContain('![Second]');
    expect(result.markdown).toContain('Between');
    expect(result.warnings.some((warning) => warning.code === 'omitted-drawing')).toBe(true);
  } finally {
    session.dispose();
  }
});
