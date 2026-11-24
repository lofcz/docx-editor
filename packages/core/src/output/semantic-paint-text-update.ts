import { kernsText } from '../layout/run-kerning.ts';
import { paragraphIsRtl } from '../layout/rtl-paragraph.ts';
import { styleForFontSlot } from '../layout/script-itemization.ts';
import type {
  LineRecord,
  ParagraphFragmentRecord,
  StyleSpanRecord,
} from '../layout/semantic-records.ts';
import { plainLineForPaint } from './plain-line-paint.ts';
import { PARAGRAPH_BOUNDARY_CLASS } from './semantic-paragraph-boundary.ts';
import { lineTerminatorEdge } from './semantic-paragraph-marks.ts';
import { setRunText } from './semantic-paint-paragraph-children.ts';
import { sameRecordExcept } from './semantic-paint-record-equality.ts';
import { interSpanGapBefore } from './semantic-paint-letter-gap.ts';

const paragraphText = new Set(['lines', 'range']);
const lineText = new Set(['id', 'spans', 'range']);
const spanText = new Set(['range', 'text', 'box', 'caretEdges']);
const lineFields = new Set([
  'id',
  'spans',
  'range',
  'box',
  'contentX',
  'baseline',
  'leading',
  'trailingSpacing',
]);
const spanFields = new Set(['range', 'text', 'box', 'props', 'style', 'fontSlot', 'caretEdges']);
const rank = () => 0;

// Keep the expected attributes from the full painter. An author slot may lend or change a node.
// The next paint uses the full reconciliation if any owned element lost its original shape.
const known = new WeakMap<
  HTMLElement,
  {
    fragment: ParagraphFragmentRecord;
    rootStyle: string;
    lineStyles: string[];
    spanStyles: string[][];
  }
>();

export function rememberTextPaint(
  root: HTMLElement,
  fresh: HTMLElement,
  fragment: ParagraphFragmentRecord
): void {
  if (fresh.children.length !== fragment.lines.length) return;
  const lines = Array.from(fresh.children) as HTMLElement[];
  if (lines.some((line) => !line.dataset.lineId)) return;
  known.set(root, {
    fragment,
    rootStyle: fresh.style.cssText,
    lineStyles: lines.map((l) => l.style.cssText),
    spanStyles: lines.map((l) => Array.from(l.children, (e) => (e as HTMLElement).style.cssText)),
  });
}

function naturalLine(line: LineRecord): boolean {
  return (
    Object.keys(line).every((key) => lineFields.has(key)) &&
    line.spans.length > 0 &&
    line.spans.every(
      (span, i) =>
        Object.keys(span).every((key) => spanFields.has(key)) &&
        !span.style.shaping &&
        span.style.horizontalScalePercent === 100 &&
        !span.style.highlight &&
        !span.style.shading &&
        span.text.length > 0 &&
        span.text.length === span.range.end - span.range.start &&
        !/[\u0000-\u001f\u007f\u00ad]/.test(span.text) &&
        interSpanGapBefore(line, i, rank) === 0
    )
  );
}

function sameTextInk(before: StyleSpanRecord, after: StyleSpanRecord): boolean {
  return (
    before.range.paragraphId === after.range.paragraphId &&
    before.box.height === after.box.height &&
    before.box.y === after.box.y &&
    sameRecordExcept(before, after, spanText) &&
    kernsText(styleForFontSlot(before.style, before.fontSlot), before.text) ===
      kernsText(styleForFontSlot(after.style, after.fontSlot), after.text)
  );
}

/** Apply text and mapping changes only when the previous complete paint proves the same ink and flow. */
export function planTextPaint(
  root: HTMLElement,
  from: ParagraphFragmentRecord,
  to: ParagraphFragmentRecord,
  scale: number
): (() => void) | null {
  const state = known.get(root);
  if (
    !state ||
    state.fragment !== from ||
    root.style.cssText !== state.rootStyle ||
    root.children.length !== to.lines.length ||
    from.lines.length !== to.lines.length ||
    from.range.start !== to.range.start ||
    paragraphIsRtl(to.props) ||
    !sameRecordExcept(from, to, paragraphText)
  )
    return null;
  const writes: (() => void)[] = [];
  for (let i = 0; i < to.lines.length; i++) {
    const before = plainLineForPaint(from.lines[i]!);
    const after = plainLineForPaint(to.lines[i]!);
    const line = root.children[i] as HTMLElement;
    const end = to.paragraphEnd && i === to.lines.length - 1;
    if (
      !naturalLine(before) ||
      !naturalLine(after) ||
      !sameRecordExcept(before, after, lineText) ||
      before.spans.length !== after.spans.length ||
      line.childNodes.length !== after.spans.length + Number(end) ||
      line.dataset.lineId !== before.id ||
      line.dataset.paragraphId !== to.paragraphId ||
      line.attributes.length !== 4 ||
      line.className !== 'docx-line layout-line' ||
      line.style.cssText !== state.lineStyles[i]
    )
      return null;
    if (before.id !== after.id)
      writes.push(() => {
        line.dataset.lineId = after.id;
      });
    for (let j = 0; j < after.spans.length; j++) {
      const old = before.spans[j]!,
        next = after.spans[j]!;
      const element = line.children[j] as HTMLElement,
        node = element?.firstChild as Text | null;
      if (
        !element ||
        element.className !== 'layout-run layout-run-text' ||
        element.attributes.length !== 5 ||
        !node ||
        node.nodeType !== 3 ||
        element.childNodes.length !== 1 ||
        node.data !== old.text ||
        element.dataset.paragraphId !== old.range.paragraphId ||
        element.dataset.start !== String(old.range.start) ||
        element.dataset.end !== String(old.range.end) ||
        element.style.cssText !== state.spanStyles[i]![j] ||
        !sameTextInk(old, next)
      )
        return null;
      if (old.range.start !== next.range.start)
        writes.push(() => {
          element.dataset.start = String(next.range.start);
        });
      if (old.range.end !== next.range.end)
        writes.push(() => {
          element.dataset.end = String(next.range.end);
        });
      if (old.text !== next.text) writes.push(() => setRunText(node, next.text));
    }
    if (end) {
      const boundary = line.lastElementChild as HTMLElement;
      if (
        boundary.className !== PARAGRAPH_BOUNDARY_CLASS ||
        boundary.childNodes.length !== 1 ||
        boundary.firstElementChild?.tagName !== 'BR' ||
        boundary.style.cssText !== state.spanStyles[i]!.at(-1)
      )
        return null;
      const x = (lineTerminatorEdge(to.lines[i]!, false).x - after.contentX) * scale;
      const previousX = (lineTerminatorEdge(from.lines[i]!, false).x - before.contentX) * scale;
      if (x !== previousX)
        writes.push(() => {
          boundary.style.left = x + 'px';
          state.spanStyles[i]![after.spans.length] = boundary.style.cssText;
        });
    }
  }
  return () => {
    for (const write of writes) write();
    state.fragment = to;
  };
}
