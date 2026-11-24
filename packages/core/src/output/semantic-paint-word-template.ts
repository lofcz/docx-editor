import type { StyleSpanRecord } from '../layout/semantic-records.ts';

const wordFields = new Set(['range', 'text', 'props', 'style', 'box', 'fontSlot', 'caretEdges']);

function naturalWord(span: StyleSpanRecord): boolean {
  return (
    span.range.end - span.range.start === span.text.length &&
    /^[\x20-\x7e\p{Script=Latin}]+$/u.test(span.text) &&
    !span.style.shaping &&
    span.style.horizontalScalePercent === 100 &&
    Object.keys(span).every((field) => wordFields.has(field))
  );
}

/**
 * Keep each word's original box and source offsets. Only its unchanged paint attributes
 * come from another word in this line. ECMA-376 Part 1 §17.3.3.31 w:t (CT_Text)
 * and §17.3.2.26 w:rFonts (CT_Fonts).
 */
export function lineWordPainter(
  paint: (span: StyleSpanRecord, band: number) => HTMLElement,
  unbound: boolean
): (span: StyleSpanRecord, band: number) => HTMLElement {
  let previous: { span: StyleSpanRecord; band: number; element: HTMLElement } | undefined;
  return (span, band) => {
    if (!naturalWord(span)) return paint(span, band);
    if (
      previous &&
      previous.span.style === span.style &&
      previous.span.props === span.props &&
      previous.span.fontSlot === span.fontSlot &&
      previous.band === band
    ) {
      const element = previous.element.cloneNode(false) as HTMLElement;
      if (!unbound) {
        element.dataset.paragraphId = span.range.paragraphId;
        element.dataset.start = String(span.range.start);
        element.dataset.end = String(span.range.end);
      }
      element.textContent = span.text;
      return element;
    }
    const element = paint(span, band);
    if (
      element.childNodes.length === 1 &&
      element.firstChild?.nodeType === 3 &&
      element.textContent === span.text
    ) {
      previous = { span, band, element: element.cloneNode(false) as HTMLElement };
    }
    return element;
  };
}
