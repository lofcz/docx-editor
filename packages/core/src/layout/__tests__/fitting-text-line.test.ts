import { describe, expect, spyOn, test } from 'bun:test';
import { readOoxmlPart } from '@docx-editor.dev/core/store';
import { DEFAULT_REVISION_MARKUP } from '../../contracts/revision-markup.ts';
import { revisionAuthorFilter } from '../revision-projection.ts';
import { literalParagraphPieces } from '../literal-paragraph-pieces.ts';
import { unmergedPiecesOfParagraphForDisplay } from '../field-projection-walk.ts';
import { breakParagraph } from '../paragraph-flow.ts';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import type { ParagraphFlowOptions } from '../paragraph-flow-options.ts';
import * as fitting from '../fitting-text-line.ts';

function paragraph(body: string) {
  const result = readOoxmlPart(
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      body +
      '</w:p>',
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!result.ok) throw Error(result.reason);
  return result.part.root;
}

function run(text: string, properties = '') {
  return (
    '<w:r>' +
    (properties ? '<w:rPr>' + properties + '</w:rPr>' : '') +
    '<w:t xml:space="preserve">' +
    text +
    '</w:t></w:r>'
  );
}

describe('literal text projection', () => {
  test('uses the same offsets, formatting and font slots as the field walk', () => {
    for (const body of [
      run('a') + run('hidden', '<w:vanish/>') + run('b'),
      '<w:pPr><w:jc w:val="right"/></w:pPr>' +
        run('中한Ａß', '<w:rFonts w:ascii="Arial" w:eastAsia="Batang"/>'),
      run('mIxEd', '<w:caps/><w:b/>') + '<w:r><w:t>ab</w:t><w:t>cd</w:t></w:r>',
    ]) {
      const input = {
        paragraph: paragraph(body),
        tocLinkStyleRanges: undefined,
        inheritedRunProperties: [{ localName: 'sz', attributes: { val: '32' } }],
      };
      expect(literalParagraphPieces(input)).toEqual(unmergedPiecesOfParagraphForDisplay(input));
    }
  });

  test('keeps visible formatting provenance when a resolved view hides a reviewer', () => {
    const change = (author: string) =>
      '<w:rPrChange w:id="3" w:author="' + author + '"><w:rPr><w:b/></w:rPr></w:rPrChange>';
    const node = paragraph(run('first', change('Alice')) + run('last', change('Bob')));
    const authorFilter = revisionAuthorFilter(['Alice']);
    for (const displayMode of ['all-markup', 'proposed', 'original'] as const) {
      const input = { paragraph: node, authorFilter, displayMode };
      expect(literalParagraphPieces(input)).toEqual(unmergedPiecesOfParagraphForDisplay(input));
      const withMarkup = {
        ...input,
        authorFilter: { ...authorFilter!, revisionMarkup: DEFAULT_REVISION_MARKUP },
      };
      if (displayMode === 'all-markup') expect(literalParagraphPieces(withMarkup)).toBeUndefined();
      else
        expect(literalParagraphPieces(withMarkup)).toEqual(
          unmergedPiecesOfParagraphForDisplay(withMarkup)
        );
    }
  });
});

describe('fitting words', () => {
  test('wraps long literal words with their measured advances, clipped spaces and fallback faces', () => {
    const base = createFixedMeasurer(5, 12);
    const fallback = createFixedMeasurer(7, 18);
    const measurer: typeof base = {
      ...base,
      measure(text, face) {
        const chosen = text.includes('Q') ? fallback : base;
        return chosen.measure(text, face) + (text.startsWith(' Q') ? -0.625 : 0);
      },
      lineMetrics(face, text) {
        return (text?.includes('Q') ? fallback : base).lineMetrics(face, text);
      },
    };
    const words = Array.from({ length: 90 }, (_, i) => `Quick ${i} map.`).join(' ');
    for (const text of [words, words + ' ' + 'x'.repeat(75) + ' end', words + ' re-entry end']) {
      const node = paragraph(run(text, '<w:kern w:val="2"/>'));
      for (const width of [52, 80.375, 147]) {
        const layout = () =>
          breakParagraph(
            node,
            node.id,
            0,
            width,
            measurer,
            undefined,
            null,
            [],
            undefined,
            undefined,
            undefined,
            { firstLineOffset: 5, contentOriginX: 3, firstLineMinimumBaseline: 17 }
          );
        const placed = layout();
        const regular = spyOn(fitting, 'placeFittingWords').mockReturnValue(undefined);
        try {
          expect(placed).toEqual(layout());
        } finally {
          regular.mockRestore();
        }
      }
    }
  });

  test('uses the same line, span and model records on both sides of a wrap', () => {
    const fixed = createFixedMeasurer(6, 14);
    let metrics = 0;
    const measurer = {
      ...fixed,
      lineMetrics(...args: Parameters<typeof fixed.lineMetrics>) {
        metrics++;
        return fixed.lineMetrics(...args);
      },
    };
    const flow: ParagraphFlowOptions = {
      firstLineOffset: 4,
      firstLineMinimumBaseline: 18,
      firstLineMarkerAscent: 12,
      contentOriginX: 3,
    };
    for (const body of [
      run('One two three four five six.'),
      run('ab-cd 0123456789-abc tail'),
      run('AB C de ', '<w:sz w:val="23"/><w:kern w:val="2"/><w:position w:val="4"/>'),
      run('coOL Text', '<w:vertAlign w:val="superscript"/><w:caps/>'),
      run('short extraordinarilylongword'),
      run('A word with a clipped space and an oversized-lettergroup near the margin.'),
    ]) {
      const node = paragraph(body);
      for (const width of [9, 28, 53, 160, 500]) {
        const layout = (options: ParagraphFlowOptions) =>
          breakParagraph(
            node,
            node.id,
            0,
            width,
            measurer,
            undefined,
            null,
            [],
            undefined,
            undefined,
            undefined,
            options
          );
        metrics = 0;
        const lines = layout(flow);
        const fittingMetrics = metrics;
        metrics = 0;
        const general = spyOn(fitting, 'placeFittingWords').mockReturnValue(undefined);
        try {
          expect(lines).toEqual(layout(flow));
        } finally {
          general.mockRestore();
        }
        if (width === 500) expect(fittingMetrics).toBeLessThan(metrics);
      }
    }
  });
});
