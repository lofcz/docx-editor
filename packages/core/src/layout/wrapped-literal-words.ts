import { automaticTextSpan } from './cjk-automatic-spacing.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import { appendWordEnd, clipWordEnd, spaceKern } from './line-end-whitespace.ts';
import { growLineMetricsForText, type PendingLine } from './pending-line.ts';
import { displayText, type ResolvedRunStyle } from './run-style.ts';
import type { TextMeasurer } from './text-measurer.ts';

interface WrappedWords {
  readonly piece: FieldAwarePiece;
  readonly face: ResolvedRunStyle;
  readonly paragraphId: string;
  readonly boundaries: readonly number[];
  readonly measurer: TextMeasurer;
  readonly tolerance: number;
  readonly nextOrigin: number;
  readonly nextAvailable: number;
  readonly nextLine: () => PendingLine;
  readonly prepare: ((boundary: number) => void) | undefined;
}

/**
 * Continue a wrapped literal at ordinary space boundaries. A word that needs character
 * breaking returns to the paragraph breaker before the line moves (ECMA-376 Part 1,
 * 17.3.1.45, w:wordWrap, CT_OnOff). The paragraph owner still finishes every line.
 */
export function placeWrappedLiteralWords(
  options: WrappedWords,
  line: PendingLine,
  first: number,
  firstWidth: number,
  origin: number,
  available: number
): number {
  const { piece, face, boundaries, measurer, tolerance } = options;
  const words = new Map<
    string,
    { width: number; metrics?: ReturnType<TextMeasurer['lineMetrics']> }
  >();
  const seams = new Map<string, number>();
  const measure = (text: string) => measurer.measure(displayText(text, face), face);
  let at = first;
  let start = boundaries[at - 1]!;
  // The final word may carry special paragraph-end spacing. The paragraph breaker owns it.
  for (; at + 1 < boundaries.length; at++) {
    const end = boundaries[at]!;
    if (piece.text[start - 1] !== ' ' || piece.text[end - 1] !== ' ') break;
    options.prepare?.(at);
    const text = piece.text.slice(start, end);
    let word = words.get(text);
    if (!word) {
      const width = at === first ? firstWidth : measure(text);
      if (!(width >= 0)) break;
      // Unique words need space only while the nearby repeated words are being placed.
      if (words.size === 256) words.clear();
      words.set(text, (word = { width }));
    }
    let clipped: ReturnType<typeof clipWordEnd>;
    if (line.width + word.width > available + tolerance) {
      clipped = clipWordEnd(text, word.width, available - line.width, measure, tolerance);
      if (!clipped) {
        if (word.width > options.nextAvailable + tolerance || line.spans.length === 0) break;
        line = options.nextLine();
        origin = options.nextOrigin;
        available = options.nextAvailable;
      }
    }
    if (line.spans.length) {
      const firstChar = text[0]!;
      let adjustment = seams.get(firstChar);
      if (adjustment === undefined) {
        adjustment = spaceKern(' ', text, face, measure);
        seams.set(firstChar, adjustment);
      }
      line.width += adjustment;
    }
    const metrics = (word.metrics ??= measurer.lineMetrics(face, displayText(text, face)));
    const width = clipped?.width ?? word.width;
    const span = automaticTextSpan(
      piece,
      { paragraphId: options.paragraphId, start, end },
      text,
      piece.style,
      origin + line.width,
      width,
      metrics.height,
      0,
      false
    );
    appendWordEnd(line.spans, span, clipped);
    line.width += width;
    line.end = end;
    line.hangingMarkBlank = undefined;
    growLineMetricsForText(line, metrics, text, face);
    start = end;
  }
  return at;
}
