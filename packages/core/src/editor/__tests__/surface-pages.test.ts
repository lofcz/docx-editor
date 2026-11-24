import { describe, expect, test } from 'bun:test';
import {
  canScrollWithoutMeasuringPages,
  retainAdjacentPages,
  surfaceExtent,
  surfacePagesInViewport,
} from '../surface-pages.ts';
import type { SemanticLayout } from '@docx-editor.dev/core/layout';

function syntheticLayout(
  pages: ReadonlyArray<{
    readonly index: number;
    readonly x?: number;
    readonly width: number;
    readonly y?: number;
    readonly height?: number;
    readonly margin?: number;
    readonly table?: { x: number; width: number };
  }>
): SemanticLayout {
  return {
    revision: 0,
    pages: pages.map((page) => ({
      index: page.index,
      box: {
        x: page.x ?? 0,
        y: page.y ?? page.index * 100,
        width: page.width,
        height: page.height ?? 100,
      },
      contentBox: { x: (page.x ?? 0) + (page.margin ?? 0), y: 0, width: page.width, height: 100 },
      fragments: page.table ? [{ kind: 'table', box: { ...page.table, y: 0, height: 20 } }] : [],
    })),
  } as unknown as SemanticLayout;
}

test('skips viewport reads only when scrolling can retain every sheet and its horizontal extent', () => {
  for (const length of [1, 2, 3, 4, 5, 20]) {
    const layout = syntheticLayout(Array.from({ length }, (_, index) => ({ index, width: 600 })));
    const built = new Set(layout.pages.map((page) => page.index));
    const alwaysRetained = layout.pages.every((page) => {
      const viewport = { top: page.box.y, height: 1 };
      const visible = surfacePagesInViewport(layout, viewport, []);
      return retainAdjacentPages(layout, visible, built)?.size === length;
    });
    expect(canScrollWithoutMeasuringPages(layout, built)).toBe(alwaysRetained);
    built.delete(0);
    expect(canScrollWithoutMeasuringPages(layout, built)).toBe(false);
  }
  const first = { index: 0, width: 600 };
  for (const second of [
    { index: 1, width: 604 },
    { index: 1, width: 600, x: 10 },
    { index: 1, width: 600, table: { x: -10, width: 650 } },
  ]) {
    expect(canScrollWithoutMeasuringPages(syntheticLayout([first, second]), new Set([0, 1]))).toBe(
      false
    );
  }
});

test('builds nearby and selected sheets without changing the horizontal scroll range', () => {
  const sheets = Array.from({ length: 5 }, (_, index) => ({
    index,
    y: index * 600,
    height: 600,
    width: 400,
  }));
  const layout = syntheticLayout(sheets);
  const start = { top: 0, height: 300 };
  expect([...surfacePagesInViewport(layout, start, [4])]).toEqual([0, 4]);
  expect([...surfacePagesInViewport(layout, { top: 160, height: 300 }, [])]).toEqual([0, 1]);
  expect([...surfacePagesInViewport(layout, { top: 650, height: 100 }, [])]).toEqual([0, 1, 2]);
  expect([...surfacePagesInViewport(layout, { top: 10000, height: 100 }, [])]).toContain(4);
  for (const next of [{ width: 700 }, { table: { x: -50, width: 700 } }]) {
    const mixed = syntheticLayout(
      sheets.map((sheet) => (sheet.index === 1 ? { ...sheet, ...next } : sheet))
    );
    expect([...surfacePagesInViewport(mixed, start, [])]).toEqual([0, 1]);
  }
});

describe('surfaceExtent width aggregation', () => {
  test('empty layout yields zero width and height', () => {
    const extent = surfaceExtent(syntheticLayout([]), undefined);
    expect(extent.width).toBe(0);
    expect(extent.height).toBe(0);
    expect(extent.pageOffsetX.size).toBe(0);
  });

  test('empty materialized set yields zero width', () => {
    const layout = syntheticLayout([{ index: 0, width: 612 }]);
    expect(surfaceExtent(layout, new Set()).width).toBe(0);
  });

  test('uses the right edge of the widest materialized page', () => {
    const layout = syntheticLayout([
      { index: 0, width: 612, x: 0 },
      { index: 1, width: 792, x: 12 },
      { index: 2, width: 500, x: 0 },
    ]);
    expect(surfaceExtent(layout, undefined).width).toBe(804);
    expect(surfaceExtent(layout, new Set([0, 2])).width).toBe(612);
    expect(surfaceExtent(layout, new Set([1])).width).toBe(804);
  });

  test('centres narrower pages when materialized widths differ', () => {
    const layout = syntheticLayout([
      { index: 0, width: 612 },
      { index: 1, width: 792 },
    ]);
    const extent = surfaceExtent(layout, new Set([0, 1]));
    expect(extent.width).toBe(792);
    expect(extent.pageOffsetX.get(0)).toBe(90);
    expect(extent.pageOffsetX.get(1)).toBe(0);
  });

  test('includes columns beyond the right paper edge only when that sheet is built', () => {
    const layout = syntheticLayout([
      { index: 0, x: 10, width: 600, margin: 36, table: { x: 12, width: 700 } },
      { index: 1, width: 600 },
    ]);
    expect(surfaceExtent(layout, new Set([0])).width).toBe(758);
    expect(surfaceExtent(layout, new Set([1])).width).toBe(600);
    expect(layout.pages[0]!.box.width).toBe(600);
  });

  test('keeps both edges reachable for a wide centred table and offsets all sheets together', () => {
    const layout = syntheticLayout([
      { index: 0, width: 600, margin: 36, table: { x: -136, width: 800 } },
      { index: 1, width: 700 },
    ]);
    const extent = surfaceExtent(layout, undefined);
    expect(extent.width).toBe(800);
    expect(extent.pageOffsetX.get(0)).toBe(100);
    expect(extent.pageOffsetX.get(1)).toBe(50);
  });

  test('handles more pages than the spread argument limit without throwing', () => {
    // V8 caps call arguments around 65k; a hostile-but-valid document can exceed that.
    const pageCount = 70_000;
    const widestIndex = 42_000;
    const widestRight = 900;
    const layout = syntheticLayout(
      Array.from({ length: pageCount }, (_, index) => ({
        index,
        width: index === widestIndex ? widestRight : 612,
      }))
    );
    expect(() => surfaceExtent(layout, undefined)).not.toThrow();
    expect(surfaceExtent(layout, undefined).width).toBe(widestRight);
    expect(surfaceExtent(layout, new Set([widestIndex])).width).toBe(widestRight);
  });
});
