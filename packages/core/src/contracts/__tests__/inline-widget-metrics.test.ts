import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import {
  MAX_INLINE_WIDGET_EXTENT,
  quantizeInlineMetric,
  resolveInlineWidgetBox,
  type InlineWidgetMetrics,
  type InlineWidgetMetricsProvider,
  type InlineWidgetRequest,
} from '../inline-widget-metrics.ts';

// Two registered inline widgets as a host would measure them beside 11pt body text: a source
// chip (citation pill with an icon) and a date chip. The numbers carry the float noise a real
// DOM measurement has, so the fixtures also pin quantization.
const SOURCE_CHIP: InlineWidgetRequest = {
  controlId: 'ctl-source',
  tag: 'scio:1:source-1',
  fallbackText: '[Komenský 1657]',
  availableWidth: 624,
};
const DATE_CHIP: InlineWidgetRequest = {
  controlId: 'ctl-date',
  tag: 'scio:1:date-1',
  fallbackText: '20. 10. 2026',
  availableWidth: 624,
};
const MEASURED: Readonly<Record<string, InlineWidgetMetrics>> = {
  'scio:1:source-1': { width: 118.4996, ascent: 15.0004, descent: 4.25 },
  'scio:1:date-1': { width: 84.0625, ascent: 13.4, descent: 3.4 },
};
const host: InlineWidgetMetricsProvider = (request) =>
  (request.tag && MEASURED[request.tag]) || null;

describe('registered inline widgets', () => {
  test('a source chip reserves its measured box on the fixed-point grid', () => {
    expect(resolveInlineWidgetBox(SOURCE_CHIP, host)).toEqual({
      kind: 'widget',
      width: 118.5,
      ascent: 15,
      descent: 4.25,
      clamped: false,
    });
  });

  test('a date chip rounds half to even, as layout rounds shaped advances', () => {
    // 84.0625 px is 84062.5 units: the even neighbour is 84062.
    expect(resolveInlineWidgetBox(DATE_CHIP, host)).toEqual({
      kind: 'widget',
      width: 84.062,
      ascent: 13.4,
      descent: 3.4,
      clamped: false,
    });
  });

  test('a chip wider than its line is clamped to the line', () => {
    const narrow = { ...SOURCE_CHIP, availableWidth: 96.0004 };
    expect(resolveInlineWidgetBox(narrow, host)).toEqual({
      kind: 'widget',
      width: 96,
      ascent: 15,
      descent: 4.25,
      clamped: true,
    });
  });

  test('a zero descent is a valid box sitting on the baseline', () => {
    const flat: InlineWidgetMetricsProvider = () => ({ width: 10, ascent: 10, descent: 0 });
    expect(resolveInlineWidgetBox(DATE_CHIP, flat)).toMatchObject({ kind: 'widget', descent: 0 });
  });
});

describe('the native path', () => {
  test('an unregistered control keeps native Word metrics', () => {
    const plain: InlineWidgetRequest = {
      controlId: 'ctl-plain',
      tag: 'customer-field',
      fallbackText: 'Field text',
      availableWidth: 624,
    };
    const untagged = { ...plain, tag: undefined };
    for (const request of [plain, untagged])
      expect(resolveInlineWidgetBox(request, host)).toEqual({
        kind: 'native',
        reason: 'unregistered',
      });
  });

  test('without a provider every control, registered or not, keeps native metrics', () => {
    for (const provider of [null, undefined])
      expect(resolveInlineWidgetBox(SOURCE_CHIP, provider)).toEqual({
        kind: 'native',
        reason: 'unregistered',
      });
  });

  test('a registered widget not measured yet reserves nothing', () => {
    const pending: InlineWidgetMetricsProvider = () => null;
    expect(resolveInlineWidgetBox(SOURCE_CHIP, pending)).toEqual({
      kind: 'native',
      reason: 'unregistered',
    });
  });

  test('measurements no box can have fall back to native metrics', () => {
    const cases: InlineWidgetMetrics[] = [
      { width: Number.NaN, ascent: 10, descent: 2 },
      { width: Number.POSITIVE_INFINITY, ascent: 10, descent: 2 },
      { width: 0, ascent: 10, descent: 2 },
      { width: -5, ascent: 10, descent: 2 },
      { width: 10, ascent: 0, descent: 2 },
      { width: 10, ascent: 10, descent: -0.5 },
      { width: MAX_INLINE_WIDGET_EXTENT + 1, ascent: 10, descent: 2 },
      { width: 10, ascent: MAX_INLINE_WIDGET_EXTENT + 1, descent: 2 },
      { width: '10', ascent: 10, descent: 2 } as unknown as InlineWidgetMetrics,
    ];
    for (const metrics of cases)
      expect(resolveInlineWidgetBox(SOURCE_CHIP, () => metrics)).toEqual({
        kind: 'native',
        reason: 'invalid-metrics',
      });
  });
});

describe('contract boundary', () => {
  test('quantization is half-to-even on the 1/1000 px grid', () => {
    expect([0.0005, 0.0015, 0.0025, 13.4, 1.23449].map(quantizeInlineMetric)).toEqual([
      0, 0.002, 0.002, 13.4, 1.234,
    ]);
  });

  test('the contract imports nothing, so no renderer or DOM reaches layout through it', () => {
    const source = readFileSync(new URL('../inline-widget-metrics.ts', import.meta.url), 'utf8');
    expect(source).not.toMatch(/^\s*import\s/m);
    expect(source).not.toMatch(/\brequire\(/);
  });
});
