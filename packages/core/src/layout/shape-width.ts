import type { ShapeInput, ShapedRun, TextShaper } from './shaped-run.ts';

export type ShapeWidthResult = ShapedRun | number;
const widths = new WeakMap<TextShaper, (input: ShapeInput) => ShapeWidthResult>();

export function registerShapeWidth(
  shaper: TextShaper,
  measure: (input: ShapeInput) => ShapeWidthResult
): void {
  widths.set(shaper, measure);
}

export function shapeForWidth(shaper: TextShaper, input: ShapeInput): ShapeWidthResult {
  const measure = widths.get(shaper);
  return measure ? measure(input) : shaper.shape(input);
}
