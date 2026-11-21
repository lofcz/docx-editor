import { describe, expect, test } from 'bun:test';
import { surfaceExtent } from '../surface-pages.ts';
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
