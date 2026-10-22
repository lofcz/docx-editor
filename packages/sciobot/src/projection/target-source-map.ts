/** Coordinate edits from one actual string normalization pass. No text diff infers these positions. */
export interface TargetNormalizationChange {
  readonly sourceStart: number;
  readonly sourceEnd: number;
  readonly targetStart: number;
  readonly targetEnd: number;
}
export interface TargetNormalizationStep {
  readonly sourceLength: number;
  readonly targetLength: number;
  readonly changes: readonly TargetNormalizationChange[];
}
export interface TargetNormalizationMap {
  readonly unit: 'utf16-code-unit';
  /** Input/output lengths compose in this order. */
  readonly steps: readonly TargetNormalizationStep[];
}
export interface TargetSourceRange {
  readonly start: number;
  readonly end: number;
  /** True only when the complete range passes through unchanged string spans. */
  readonly exact: boolean;
  /** The range crosses syntax inserted without original source characters. */
  readonly generated: boolean;
}
export interface NormalizationEdit {
  readonly start: number;
  readonly end: number;
  readonly text: string;
}

/** Called only with edits captured by the string normalizers, before AST transforms run. */
export function normalizationStep(
  sourceLength: number,
  targetLength: number,
  edits: readonly NormalizationEdit[]
): TargetNormalizationStep {
  let delta = 0,
    previous = 0;
  const changes = [...edits]
    .sort((a, b) => a.start - b.start)
    .map((edit) => {
      if (edit.start < previous || edit.end < edit.start || edit.end > sourceLength)
        throw new TypeError('Normalization edits overlap or exceed the source.');
      const targetStart = edit.start + delta,
        targetEnd = targetStart + edit.text.length;
      delta += edit.text.length - (edit.end - edit.start);
      previous = edit.end;
      return Object.freeze({
        sourceStart: edit.start,
        sourceEnd: edit.end,
        targetStart,
        targetEnd,
      });
    });
  if (sourceLength + delta !== targetLength)
    throw new TypeError('Normalization edits do not explain the target length.');
  return Object.freeze({ sourceLength, targetLength, changes: Object.freeze(changes) });
}

function point(step: TargetNormalizationStep, offset: number): number | null {
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > step.targetLength) return null;
  const changes = step.changes;
  let low = 0,
    high = changes.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (changes[middle]!.targetStart <= offset) low = middle + 1;
    else high = middle;
  }
  const change = changes[low - 1];
  if (!change) return offset;
  if (offset === change.targetStart) return change.sourceStart;
  if (offset < change.targetEnd) return null;
  return change.sourceEnd + offset - change.targetEnd;
}

/** Read a parser-owned trace. This maps string coordinates, not native text ownership or writability. */
export function normalizationSourceRangeAt(
  map: TargetNormalizationMap,
  start: number,
  end: number
): TargetSourceRange | null {
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start)
    return null;
  let exact = true,
    generated = false;
  for (let index = map.steps.length - 1; index >= 0; index--) {
    const step = map.steps[index]!,
      from = point(step, start),
      to = point(step, end);
    if (from === null || to === null) return null;
    let low = 0,
      high = step.changes.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (step.changes[middle]!.targetEnd <= start) low = middle + 1;
      else high = middle;
    }
    for (let i = low; i < step.changes.length && step.changes[i]!.targetStart < end; i++) {
      const change = step.changes[i]!;
      exact = false;
      if (change.sourceStart === change.sourceEnd) generated = true;
    }
    start = from;
    end = to;
  }
  return Object.freeze({ start, end, exact, generated });
}

export function splitsSurrogate(text: string, offset: number): boolean {
  const before = text.charCodeAt(offset - 1),
    after = text.charCodeAt(offset);
  return before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff;
}
