import { geometryOfSection } from '../layout/index.ts';
import { verticalMarginInsets } from '../layout/page-body-margins.ts';
import { lineAtIndexedPosition } from '../layout/paragraph-lines.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';
import type { SelectedImageState } from '../contracts/editor.ts';

export function naturalExtentEmu(
  image: SelectedImageState
): { readonly cx: number; readonly cy: number } | null {
  if (!image.intrinsic) return null;
  const cx = Math.round((image.intrinsic.pixelWidth * 914_400) / image.intrinsic.dpiX);
  const cy = Math.round((image.intrinsic.pixelHeight * 914_400) / image.intrinsic.dpiY);
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) return null;
  return { cx, cy };
}

/**
 * Scale a natural extent down to fit where the caret flows, preserving aspect ratio —
 * Word's insert behavior: an image that fits keeps its size, a wider or taller one lands at
 * the width of its cell, column, or page. Only ever shrinks.
 *
 * Width fits the caret LINE's measure, not the section box alone: a table cell or a text
 * column is narrower than the page, and layout hard-clips an inline drawing that overflows
 * its box. `geometryOfSection` carries layout's own degenerate-margins fallback, so a page
 * whose margins swallow it still fits against the box it actually paginates into.
 */
export function fitInsertExtent(
  surface: PaginatedSurface,
  paragraphId: string,
  offset: number,
  widthPoints: number,
  heightPoints: number
): { readonly widthPoints: number; readonly heightPoints: number } {
  const geometry = geometryOfSection(surface.sectionPropertiesAt(paragraphId));
  const contentWidth = geometry.width - geometry.margin.left - geometry.margin.right;
  const margins = verticalMarginInsets(geometry);
  const contentHeight = geometry.height - margins.top - margins.bottom;
  const line = lineAtIndexedPosition(surface.layout(), paragraphId, offset);
  const flowWidth =
    line && line.box.width > 0 ? Math.min(line.box.width, contentWidth) : contentWidth;
  const scale = Math.min(1, flowWidth / widthPoints, contentHeight / heightPoints);
  if (scale >= 1) return { widthPoints, heightPoints };
  // ONE shared scale, unrounded: independent per-axis rounding or floors distort an
  // extreme-aspect image (a hairline rule, a tall banner).
  return { widthPoints: widthPoints * scale, heightPoints: heightPoints * scale };
}
