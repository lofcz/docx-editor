import type { ShapeInput, TextShaper } from './shaped-run.ts';

/** Internal capability. Custom shapers always keep their own complete run semantics. */
export const measureShapedWidth = Symbol('measureShapedWidth');
type WidthShaper = TextShaper & {
  [measureShapedWidth]?(input: ShapeInput): number | undefined;
};

export function widthProbeOf(shaper: TextShaper) {
  return (shaper as WidthShaper)[measureShapedWidth]?.bind(shaper);
}
