import { expect, test } from 'bun:test';
import {
  freezeMarkdownSourceMap,
  preserveLeadingWhitespace,
  quoteMarkdownLines,
  escapeUnescapedTablePipes,
} from '../src/markdown-source-map.ts';

test('keeps leading whitespace one-to-one without emitting HTML entities or code indentation', () => {
  const value = preserveLeadingWhitespace({
    markdown: ' \tText',
    sources: [
      {
        sourceScope: 'body',
        paragraphId: 'p',
        sourceStart: 0,
        sourceEnd: 6,
        markdownStart: 0,
        markdownEnd: 6,
        markdownBoundaries: [
          { sourceOffset: 0, markdownOffset: 0 },
          { sourceOffset: 2, markdownOffset: 2 },
          { sourceOffset: 6, markdownOffset: 6 },
        ],
        exact: true,
      },
    ],
  });

  expect(value.markdown).toBe('\u00a0\u00a0Text');
  expect(value.markdown).not.toContain('&nbsp;');
  expect(value.sources[0]).toMatchObject({
    markdownStart: 0,
    markdownEnd: 6,
    markdownBoundaries: [
      { sourceOffset: 0, markdownOffset: 0 },
      { sourceOffset: 2, markdownOffset: 2 },
      { sourceOffset: 6, markdownOffset: 6 },
    ],
  });
});

test('keeps complete native boundaries around generated line and table syntax', () => {
  const source = {
    markdown: 'a\nb|c',
    sources: [
      {
        sourceScope: 'body',
        paragraphId: 'p',
        sourceStart: 0,
        sourceEnd: 5,
        markdownStart: 0,
        markdownEnd: 5,
        exact: true,
        completeBoundaries: true,
        markdownBoundaries: [
          { sourceOffset: 0, markdownOffset: 0 },
          { sourceOffset: 5, markdownOffset: 5 },
        ],
      },
    ],
  };
  const transformed = escapeUnescapedTablePipes(quoteMarkdownLines(source));
  expect(transformed.markdown).toBe('> a\n> b\\|c');
  expect(freezeMarkdownSourceMap(transformed).sources[0]?.markdownBoundaries).toEqual([
    { sourceOffset: 0, markdownOffset: 2 },
    { sourceOffset: 1, markdownOffset: 3 },
    { sourceOffset: 2, markdownOffset: 6 },
    { sourceOffset: 3, markdownOffset: 7 },
    { sourceOffset: 4, markdownOffset: 9 },
    { sourceOffset: 5, markdownOffset: 10 },
  ]);
});
