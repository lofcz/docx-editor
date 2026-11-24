import { displayText, type ResolvedRunStyle } from './run-style.ts';
import type { TextMeasurer } from './text-measurer.ts';

type Prepare = (
  text: string,
  boundaries: readonly number[],
  first: number,
  end: number,
  style: ResolvedRunStyle
) => void;
const preparations = new WeakMap<TextMeasurer, Prepare>();
const WIDTH_WINDOW = 64;
const MAX_PREPARATION_UNITS = 512;

export function registerWidthPreparation(measurer: TextMeasurer, prepare: Prepare): void {
  preparations.set(measurer, prepare);
}

export function forwardWidthPreparation(from: TextMeasurer, to: TextMeasurer): void {
  const prepare = preparations.get(from);
  if (prepare) preparations.set(to, prepare);
}

export function pieceWidthPreparation(
  measurer: TextMeasurer,
  text: string,
  boundaries: readonly number[],
  style: ResolvedRunStyle
) {
  const prepare = preparations.get(measurer);
  if (
    !prepare ||
    boundaries.length < WIDTH_WINDOW ||
    style.shaping?.context ||
    style.shaping?.direction === 'rtl'
  )
    return undefined;
  let next = 0;
  return (first: number) => {
    if (first < next) return;
    next = Math.min(first + WIDTH_WINDOW, boundaries.length);
    prepare(text, boundaries, first, next, style);
  };
}

/** Only prepare printable Latin text. Other scripts keep their normal shaping context and feature ranges. */
export function missingWidthParts(
  text: string,
  boundaries: readonly number[],
  first: number,
  end: number,
  style: ResolvedRunStyle,
  known: ReadonlyMap<string, number>
) {
  const missing = new Set<string>();
  let length = 0;
  for (let i = first; i < end; i++) {
    const start = i ? boundaries[i - 1]! : 0;
    const stop = boundaries[i]!;
    if (stop - start > MAX_PREPARATION_UNITS) continue;
    const part = displayText(text.slice(start, stop), style);
    if (!part || known.has(part) || missing.has(part) || /[^ -~]/u.test(part)) continue;
    if (length + part.length > MAX_PREPARATION_UNITS) break;
    missing.add(part);
    length += part.length;
  }
  return missing;
}
