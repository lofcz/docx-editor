import { paragraphFragmentsOf, type SemanticLayout } from '@docx-editor.dev/core/layout';

interface InlineBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}
const cache = new WeakMap<SemanticLayout, ReadonlyMap<string, InlineBox>>();

/** Native widget extents, rather than the surrounding paragraph's selection band. */
export function inlineLayoutBoxes(layout: SemanticLayout): ReadonlyMap<string, InlineBox> {
  const previous = cache.get(layout);
  if (previous) return previous;
  const boxes = new Map<string, InlineBox>();
  for (const page of layout.pages)
    for (const paragraph of paragraphFragmentsOf(page))
      for (const line of paragraph.lines)
        for (const span of line.spans) {
          const widget = span.inlineWidget;
          if (widget)
            boxes.set(
              `${page.index}:${widget.controlId}`,
              Object.freeze({
                x: span.box.x,
                y: line.box.y + line.baseline - widget.ascent,
                width: widget.width,
                height: widget.ascent + widget.descent,
              })
            );
        }
  cache.set(layout, boxes);
  return boxes;
}
