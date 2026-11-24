import { automaticTextSpan } from './cjk-automatic-spacing.ts';
import { wordBoundaries } from './cjk-line-break.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import { fitTextGroupOf } from './fit-text-groups.ts';
import { spaceKern } from './line-end-whitespace.ts';
import { growLineMetricsForText, type PendingLine } from './pending-line.ts';
import { displayText } from './run-style.ts';
import { styleForFontSlot } from './script-itemization.ts';
import type { TextMeasurer } from './text-measurer.ts';

// ECMA-376 Part 1, 17.3.3.31, w:t (CT_Text): keep literal text and its UTF-16 range.
// Admit basic printable text and Latin letters. Other scripts and combining marks use the general breaker.
const needsGeneralText = /[^ -~\p{Script=Latin}]|^ | {2}/u;

const textFields = new Set(['text', 'props', 'style', 'start', 'end', 'link', 'fontSlot']);

/**
 * Place one literal run while every word fits the unbroken line. Each word keeps its own
 * measured advance and model range. At the first overflow the paragraph breaker decides
 * clipping, shrink, hanging punctuation, and wrapping against the actual margin.
 */
export function placeFittingWords(
  line: PendingLine,
  piece: FieldAwarePiece,
  paragraphId: string,
  origin: number,
  available: number,
  measurer: TextMeasurer,
  tolerance: number
): { readonly boundaries: readonly number[]; readonly overflowWidth?: number } | undefined {
  if (
    piece.start !== 0 ||
    piece.end !== piece.text.length ||
    !piece.text.length ||
    needsGeneralText.test(piece.text) ||
    piece.style.shaping?.context ||
    fitTextGroupOf(piece) !== undefined ||
    Object.keys(piece).some(
      (key) => !textFields.has(key) && piece[key as keyof FieldAwarePiece] !== undefined
    )
  )
    return undefined;
  const face = styleForFontSlot(piece.style, piece.fontSlot);
  const measure = (text: string) => measurer.measure(displayText(text, face), face);
  let start = 0;
  let previous = '';
  const boundaries = wordBoundaries(piece.text, false);
  for (const end of boundaries) {
    const text = piece.text.slice(start, end);
    const displayed = displayText(text, face);
    const width = measurer.measure(displayed, face);
    if (!(width >= 0 && line.width + width <= available + tolerance)) {
      return { boundaries, overflowWidth: width };
    }
    line.width += spaceKern(previous, text, face, measure);
    const metrics = measurer.lineMetrics(face, displayed);
    line.spans.push(
      automaticTextSpan(
        piece,
        { paragraphId, start, end },
        text,
        piece.style,
        origin + line.width,
        width,
        metrics.height,
        0,
        false
      )
    );
    line.width += width;
    line.end = end;
    line.hangingMarkBlank = undefined;
    growLineMetricsForText(line, metrics, text, face);
    previous = text;
    start = end;
  }
  return { boundaries };
}
