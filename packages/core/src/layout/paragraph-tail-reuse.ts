// Reuse the rest of a sheet when a changed text paragraph occupies the same place.
// Store the new paragraph in the page and in every checkpoint that can resume from it.
import type { FlowCheckpoint } from './layout-session.ts';
import type {
  LineRecord,
  PageRecord,
  ParagraphFragmentRecord,
  StyleSpanRecord,
} from './semantic-records.ts';
import { flowCheckpointsMatch } from './flow-checkpoint.ts';
import { sameFragments } from './semantic-fragment-signature.ts';
import { bodyLineId } from './body-line-id.ts';

const TEXT_FIELDS = new Set([
  'id',
  'range',
  'text',
  'props',
  'style',
  'fontSlot',
  'syntheticBold',
  'box',
  'lineEndWhitespace',
  'caretEdges',
  'borderBaselinePt',
]);

function plain(span: StyleSpanRecord): boolean {
  return Object.keys(span).every((key) => TEXT_FIELDS.has(key)) && !/[\t\r\n\v\f]/.test(span.text);
}

// New layout fields must opt into this fast path before a paragraph can carry them.
const TEXT_LINE_FIELDS = new Set([
  'id',
  'range',
  'spans',
  'box',
  'contentX',
  'baseline',
  'leading',
  'trailingSpacing',
]);

function ordinaryTextLine(line: LineRecord): boolean {
  return (
    Object.keys(line).every((key) => TEXT_LINE_FIELDS.has(key)) &&
    line.spans.length > 0 &&
    line.spans.every(plain)
  );
}

function sameTextFlow(old: ParagraphFragmentRecord, next: ParagraphFragmentRecord): boolean {
  if (
    !old.paragraphEnd ||
    old.fragmentIndex ||
    old.outOfFlow ||
    old.localPlaceholder ||
    old.positionedFrame ||
    old.range.start !== 0 ||
    next.range.start !== 0 ||
    old.range.paragraphId !== next.range.paragraphId ||
    old.lines.length !== next.lines.length
  )
    return false;
  if (!sameFragments([old], [{ ...next, lines: old.lines, range: old.range }])) return false;
  return old.lines.every((line, i) => {
    const after = next.lines[i]!;
    // A w:p (CT_P, ECMA-376 Part 1 §17.3.1.22) owns these replacement lines.
    // Typing can move a later line's start offset without moving its painted box.
    const sameLine =
      line.id === after.id ||
      (line.id === bodyLineId(line.range.paragraphId, line.range.start, i) &&
        after.id === bodyLineId(after.range.paragraphId, after.range.start, i));
    if (
      !sameLine ||
      line.range.paragraphId !== after.range.paragraphId ||
      !ordinaryTextLine(line) ||
      !ordinaryTextLine(after)
    )
      return false;
    return (
      line.box.x === after.box.x &&
      line.box.y === after.box.y &&
      line.box.width === after.box.width &&
      line.box.height === after.box.height &&
      line.baseline === after.baseline &&
      line.leading === after.leading &&
      line.trailingSpacing === after.trailingSpacing
    );
  });
}

export function reuseAfterText(
  prior: readonly FlowCheckpoint[],
  index: number,
  current: FlowCheckpoint,
  page: PageRecord | undefined
) {
  const old = prior[index]!;
  if (
    !page ||
    old.pageCount !== current.pageCount ||
    old.lineCounter !== current.lineCounter ||
    old.pageFragments.length !== current.pageFragments.length ||
    page.hasBodyPageFields ||
    page.footnotes ||
    page.endnotes ||
    page.anchoredDrawings?.length ||
    !flowCheckpointsMatch(old, { ...current, pageFragments: old.pageFragments })
  )
    return null;
  let changed = -1;
  for (let i = 0; i < old.pageFragments.length; i++) {
    const before = old.pageFragments[i]!;
    const after = current.pageFragments[i]!;
    if (sameFragments([before], [after])) continue;
    if (
      changed >= 0 ||
      before.kind !== 'paragraph' ||
      after.kind !== 'paragraph' ||
      !sameTextFlow(before, after) ||
      page.fragments[i] !== before
    )
      return null;
    changed = i;
  }
  if (changed < 0) return null;
  const from = old.pageFragments[changed]!;
  const to = current.pageFragments[changed]!;
  let end = index;
  while (end < prior.length && prior[end]!.pageCount === old.pageCount) {
    if (prior[end]!.pageFragments[changed] !== from) return null;
    end++;
  }
  return {
    page: {
      ...page,
      fragments: page.fragments.map((fragment) => (fragment === from ? to : fragment)),
    },
    updateCheckpoints(marks: FlowCheckpoint[]) {
      for (let at = index; at < end; at++) {
        const mark = marks[at]!;
        const fragments = mark.pageFragments.slice();
        fragments[changed] = to;
        marks[at] = { ...mark, pageFragments: fragments };
      }
    },
  };
}
