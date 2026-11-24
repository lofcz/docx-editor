import { setCharacterFillBand } from './character-fill-band.ts';
import { inlineExtentSpacing, lineSpacingBand } from './equation-line-spacing.ts';
import type { ParagraphMarkStyles } from './paragraph-mark-styles.ts';
import { growScriptLineMetrics } from './paragraph-mark-metrics.ts';
import { markSyntheticBoldSpans } from './paragraph-span-metadata.ts';
import type { ParagraphLineSpacing } from './paragraph-style.ts';
import { repositionPendingDrawings, syncPendingDrawingBaselines } from './pending-line-drawings.ts';
import {
  clipLineDrawingsAtTop,
  growPendingLineDrawingExtent,
  isHeightlessWhitespace,
  type PendingLine,
} from './pending-line.ts';
import { spacedLine } from './positioned-run-spacing.ts';
import { growRunBorderLineMetrics } from './run-border-strokes.ts';
import type { TextMeasurer } from './text-measurer.ts';

/** Apply the paragraph mark, run, drawing, list, and spacing metrics to the finished line. */
export function finishParagraphLineMetrics(
  line: PendingLine,
  styles: ParagraphMarkStyles,
  measurer: TextMeasurer,
  lineSpacing: ParagraphLineSpacing,
  firstLine: boolean,
  firstLineOffset: number,
  markerAscent: number,
  markerBaselineFloor: number,
  pageBreaksIgnored: boolean,
  pictureRunBand?: { readonly height: number }
): boolean {
  const empty =
    line.drawings.length === 0 && line.spans.every((span) => isHeightlessWhitespace(span.text));
  const metrics = measurer.lineMetrics(empty ? styles.empty : styles.cascade);
  // Baseline of the visible glyph band before mark / spacing. Paint's padding-top is
  // `spaced.baseline - glyphBaseline` (space above); auto extras grow BELOW instead.
  let glyphBaseline = line.baseline;
  if (line.height === 0) {
    line.height = metrics.height;
    line.baseline = metrics.baseline;
    glyphBaseline = metrics.baseline;
  } else {
    glyphBaseline += growScriptLineMetrics(line, styles.unstyledMark.verticalAlign, measurer);
  }
  // The list marker is painted as furniture, but it sits on THIS line's baseline, so its
  // face reserves space above it like the run the marker is in Word. The descent is the
  // text's alone. Only the paragraph's first line carries a marker.
  if (firstLine && firstLineOffset !== 0) line.firstLineOffset = firstLineOffset;
  if (firstLine && markerAscent > line.baseline) {
    const raised = markerAscent - line.baseline;
    line.baseline = markerAscent;
    line.height += raised;
    glyphBaseline += raised;
  }
  growRunBorderLineMetrics(line, measurer);
  markSyntheticBoldSpans(line, measurer);
  syncPendingDrawingBaselines(line, pageBreaksIgnored);
  growPendingLineDrawingExtent(line);
  // Apply paragraph line spacing once to the finished box.
  const naturalHeight = line.height;
  const naturalBaseline = line.baseline;
  // `auto` scales the text band only beside an inline drawing or equation (ECMA-376 17.3.1.33).
  const { gridFitsDrawing, scalesTextBandOnly, addsTextBandExtra, removesTextBand } =
    inlineExtentSpacing(line, lineSpacing);
  const fallbackBand = pictureRunBand?.height ?? metrics.height;
  const band = scalesTextBandOnly
    ? lineSpacingBand(line, lineSpacing, measurer, fallbackBand, pageBreaksIgnored)
    : undefined;
  const equationBand = band?.equation ? band : undefined;
  const spacingBase = band?.height ?? naturalHeight;
  const bandBaseline = band?.baseline ?? line.baseline;
  const spaced = spacedLine(lineSpacing, spacingBase, bandBaseline, line.spans, measurer);
  if (!scalesTextBandOnly) line.baseline = spaced.baseline;
  else if (equationBand) line.baseline += spaced.baseline - bandBaseline;
  const floored = firstLine && lineSpacing.rule !== 'exact' ? markerBaselineFloor : 0;
  const markerFloor = Math.max(0, floored - line.baseline);
  line.baseline += markerFloor;
  // Space ABOVE the glyph band only (exact baseline placement, not auto/atLeast). Never negative.
  line.leading = Math.max(0, line.baseline - glyphBaseline);
  const textBandExtra = equationBand
    ? spaced.height - spacingBase
    : addsTextBandExtra
      ? Math.max(0, spaced.height - spacingBase)
      : removesTextBand
        ? Math.min(0, spaced.height - spacingBase)
        : 0;
  line.height = scalesTextBandOnly
    ? Math.max(spaced.height, naturalHeight + textBandExtra)
    : spaced.height;
  // A baseline below the shortened line's foot moves up to it, so a picture ends there.
  if (removesTextBand) line.baseline = Math.min(line.baseline, line.height);
  // Like a text line's auto extra, the depth below the drawing may cross the bottom margin.
  const drawingLineTrailing = addsTextBandExtra ? Math.max(0, line.height - naturalHeight) : 0;
  const spacedExtent = equationBand ? line.height : spaced.height;
  line.height += markerFloor;
  // Baseline shifts from line spacing must move inline drawings too, or authored distT/distB
  // and the text baseline drift apart. For `exact`, keep the authored box — tall drawings
  // clip/overflow per content-clip policy; auto/atLeast still grow to contain distB.
  repositionPendingDrawings(line);
  // A picture taller than a shortened line paints only below the line's top.
  if (removesTextBand) clipLineDrawingsAtTop(line);
  if (lineSpacing.rule !== 'exact') growPendingLineDrawingExtent(line);
  // A grid line's centring below its content may hang past the page foot, picture or not.
  line.trailingSpacing =
    (line.drawings.length === 0 || gridFitsDrawing) && lineSpacing.rule !== 'exact'
      ? Math.max(0, Math.min(spaced.trailing ?? Infinity, spacedExtent - naturalHeight))
      : drawingLineTrailing;
  setCharacterFillBand(line, lineSpacing.rule === 'exact', {
    height: naturalHeight,
    baseline: naturalBaseline,
  });
  return empty;
}
