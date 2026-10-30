import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPart,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { createFixedMeasurer } from '../fixed-measurer.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { linesOf, type TextMeasurer } from '../semantic-records.ts';
import { caretAt } from '../semantic-interaction.ts';
import { hitTestPage } from '../semantic-hit-test.ts';
import { createLayoutSession } from '../layout-session.ts';
import type { PendingLine } from '../pending-line.ts';
import { createParagraphLayoutCache } from '../layout-cache.ts';
import { paintSemanticLayout } from '../../output/semantic-paint.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r>`;
// Word shows boundary spaces in w:t only with xml:space="preserve".
const preserved = (xml: string) => xml.replaceAll('<w:t>', '<w:t xml:space="preserve">');
const chip = `<w:sdt><w:sdtPr><w:tag w:val="chip:source"/></w:sdtPr><w:sdtContent>${run('source')}</w:sdtContent></w:sdt>`;
function part(body = `<w:p>${run('A')}${chip}${run('B')}</w:p>`) {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
const geometry = { width: 100, height: 200, margin: { top: 10, bottom: 10, left: 10, right: 10 } };
const measure = (width = 40, revision = '1'): TextMeasurer => ({
  ...createFixedMeasurer(),
  inlineWidgets: {
    measure: (request) =>
      request.tag === 'chip:source' ? { width, ascent: 20, descent: 4 } : null,
    revision,
  },
});

test('CSS measurements reserve exact point advances and adjacent native caret coordinates without source writes', () => {
  const source = part(),
    before = serializeOoxmlPart(source);
  const layout = layoutSemanticDocument(source, 1, { measurer: measure(), geometry });
  const lines = linesOf(layout),
    spans = lines.flatMap((line) => line.spans);
  expect(lines).toHaveLength(1);
  expect(spans.map((span) => [span.text, span.box.x, span.box.width])).toEqual([
    ['A', 0, 6],
    ['source', 6, 30],
    ['B', 36, 6],
  ]);
  const widget = spans[1]!;
  expect(widget.inlineWidget).toMatchObject({ width: 30, ascent: 15, descent: 3, clamped: false });
  expect(widget.range).toMatchObject({ start: 1, end: 7 });
  expect(lines[0]!.baseline).toBe(15);
  expect(lines[0]!.box.height).toBe(18);
  expect(caretAt(layout, { paragraphId: widget.range.paragraphId, offset: 1 })?.x).toBe(6);
  expect(caretAt(layout, { paragraphId: widget.range.paragraphId, offset: 7 })?.x).toBe(36);
  expect(hitTestPage(layout, 0, { x: 8, y: 5 })?.position.offset).toBe(1);
  expect(hitTestPage(layout, 0, { x: 34, y: 5 })?.position.offset).toBe(7);
  expect(serializeOoxmlPart(source)).toBe(before);
  expect(paragraphTextOf(source, widget.range.paragraphId)).toBe('AsourceB');
});

test('a chip wraps indivisibly at its measured width and clamps an oversized answer', () => {
  const source = part(`<w:p>${run('AA')}${chip}${run('Z')}</w:p>`);
  const layout = layoutSemanticDocument(source, 1, { measurer: measure(96), geometry });
  const lines = linesOf(layout);
  expect(lines.map((line) => line.spans.map((span) => span.text))).toEqual([
    ['AA'],
    ['source', 'Z'],
  ]);
  expect(lines[1]!.spans.map((span) => [span.box.x, span.box.width])).toEqual([
    [0, 72],
    [72, 6],
  ]);
  const clamped = linesOf(layoutSemanticDocument(source, 1, { measurer: measure(200), geometry }))
    .flatMap((line) => line.spans)
    .find((span) => span.inlineWidget)!;
  expect(clamped.inlineWidget).toMatchObject({ width: 80, clamped: true });
});

test('missing, declined, and invalid widget metrics keep native Word fallback layout', () => {
  const source = part();
  const baseline = layoutSemanticDocument(source, 1, { measurer: createFixedMeasurer(), geometry });
  for (const provider of [
    () => null,
    () => ({ width: NaN, ascent: 20, descent: 4 }),
    () => ({ width: 40, ascent: 0, descent: 4 }),
  ]) {
    const layout = layoutSemanticDocument(source, 1, {
      measurer: { ...createFixedMeasurer(), inlineWidgets: { measure: provider, revision: '1' } },
      geometry,
    });
    expect(linesOf(layout)).toEqual(linesOf(baseline));
  }
});

test('measurement revision and provider identity invalidate warm paragraph and incremental layout caches', () => {
  const source = part(),
    cache = createParagraphLayoutCache<readonly PendingLine[]>();
  let width = 40;
  const provider = () => ({ width, ascent: 20, descent: 4 });
  const options = {
    geometry,
    cache,
    session: createLayoutSession(),
    producer: 'same-fonts',
    measurer: { ...createFixedMeasurer(), inlineWidgets: { measure: provider, revision: '1' } },
  };
  const first = layoutSemanticDocument(source, 1, options);
  width = 64;
  const second = layoutSemanticDocument(source, 1, {
    ...options,
    measurer: { ...options.measurer, inlineWidgets: { measure: provider, revision: '2' } },
  });
  const widget = (layout: typeof first) =>
    linesOf(layout)
      .flatMap((line) => line.spans)
      .find((span) => span.inlineWidget)!;
  expect(widget(first).box.width).toBe(30);
  expect(widget(second).box.width).toBe(48);
  const third = layoutSemanticDocument(source, 1, {
    ...options,
    measurer: {
      ...options.measurer,
      inlineWidgets: { measure: () => ({ width: 80, ascent: 20, descent: 4 }), revision: '2' },
    },
  });
  expect(widget(third).box.width).toBe(60);
});

test('paint reserves the same inline advance while native fallback and source fingerprints remain intact', () => {
  const source = part(),
    before = canonicalOoxmlFingerprint(source.root);
  const layout = layoutSemanticDocument(source, 1, { measurer: measure(), geometry });
  const host = document.createElement('div');
  paintSemanticLayout(host, layout, { scale: 2 });
  const element = host.querySelector<HTMLElement>('[data-docx-inline-widget]')!;
  expect(element).not.toBeNull();
  expect(element.style.width).toBe('60px');
  expect(element.textContent).toBe('source');
  expect(element.contentEditable).toBe('false');
  expect(canonicalOoxmlFingerprint(source.root)).toBe(before);
});

test('RTL itemization and justification keep a measured whitespace-bearing chip whole', () => {
  for (const properties of ['<w:bidi/>', '<w:jc w:val="both"/>']) {
    const source = part(
      `<w:p><w:pPr>${properties}</w:pPr>${run('אב')}${preserved(chip.replace('>source<', '>source <'))}${preserved(run(' גד words words words'))}</w:p>`
    );
    const measured = layoutSemanticDocument(source, 1, { measurer: measure(40), geometry });
    const widgets = linesOf(measured)
      .flatMap((line) => line.spans)
      .filter((span) => span.inlineWidget);
    expect(widgets).toHaveLength(1);
    expect(widgets[0]!.box.width).toBe(30);
    expect(widgets[0]!.text).toBe('source ');
    expect(widgets[0]!.range.end - widgets[0]!.range.start).toBe(7);
  }
});

test('registered and declined inline controls keep their own geometry inside a table cell', () => {
  const source = part(
    `<w:tbl><w:tblGrid><w:gridCol w:w="1600"/></w:tblGrid><w:tr><w:tc><w:p>${run('A')}${chip}${run('B')}</w:p></w:tc></w:tr></w:tbl>`
  );
  const measured = layoutSemanticDocument(source, 1, { measurer: measure(), geometry });
  const widget = linesOf(measured)
    .flatMap((line) => line.spans)
    .find((span) => span.inlineWidget)!;
  expect(widget).toBeDefined();
  expect(widget.box.width).toBe(30);
  expect(widget.range).toMatchObject({ start: 1, end: 7 });
  const declined: TextMeasurer = {
    ...createFixedMeasurer(),
    inlineWidgets: { measure: () => null, revision: '1' },
  };
  const plain = layoutSemanticDocument(source, 1, { measurer: createFixedMeasurer(), geometry });
  expect(linesOf(layoutSemanticDocument(source, 1, { measurer: declined, geometry }))).toEqual(
    linesOf(plain)
  );
});

test('mixed native revisions retain exact native spans in every review view instead of losing attribution in a widget', () => {
  const source = part(
    `<w:p>${run('A')}<w:sdt><w:sdtPr><w:tag w:val="chip:source"/></w:sdtPr><w:sdtContent>${run('old')}<w:ins w:id="17" w:author="Peer">${run('new')}</w:ins></w:sdtContent></w:sdt>${run('B')}</w:p>`
  );
  const before = canonicalOoxmlFingerprint(source.root);
  for (const displayMode of ['all-markup', 'original', 'proposed'] as const) {
    const plain = layoutSemanticDocument(source, 1, {
      measurer: createFixedMeasurer(),
      geometry,
      displayMode,
    });
    const measured = layoutSemanticDocument(source, 1, {
      measurer: measure(),
      geometry,
      displayMode,
    });
    expect(linesOf(measured)).toEqual(linesOf(plain));
    expect(
      linesOf(measured)
        .flatMap((line) => line.spans)
        .some((span) => span.inlineWidget)
    ).toBe(false);
  }
  expect(canonicalOoxmlFingerprint(source.root)).toBe(before);
});
