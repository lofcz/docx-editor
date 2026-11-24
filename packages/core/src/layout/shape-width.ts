import type { ShapeInput, ShapedRun, TextShaper } from './shaped-run.ts';

export type ShapeWidthResult = ShapedRun | number;
const widths = new WeakMap<TextShaper, (input: ShapeInput) => ShapeWidthResult>();
type PartitionWidths = (
  input: ShapeInput,
  ends: readonly number[]
) => readonly number[] | undefined;
const partitions = new WeakMap<TextShaper, PartitionWidths>();

export function registerShapeWidth(
  shaper: TextShaper,
  measure: (input: ShapeInput) => ShapeWidthResult,
  partition?: PartitionWidths
): void {
  widths.set(shaper, measure);
  if (partition) partitions.set(shaper, partition);
}

export function shapeForWidth(shaper: TextShaper, input: ShapeInput): ShapeWidthResult {
  const measure = widths.get(shaper);
  return measure ? measure(input) : shaper.shape(input);
}

export function partitionWidthsOf(shaper: TextShaper): PartitionWidths | undefined {
  return partitions.get(shaper);
}
