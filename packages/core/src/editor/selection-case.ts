import type { ExecResult, ParagraphText } from '../contracts/editor.ts';
import type { TreeDocOp } from '@docx-editor.dev/core/store';
import { findNode } from '../store/package/ooxml-edit.ts';
import { segmentsOf } from '../store/store/tree-op-segments.ts';
import { caseLetters, type CaseMode } from './selection-case-text.ts';
import { matchRangeSkip } from './docx-editor-replace-all-ranges.ts';
import { acceptedOffsetFromText } from './paragraph-text-ranges.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';
import { partOfNodeId } from './surface-scope.ts';
import { openTextboxParagraphOrder } from './review-story-entry.ts';
import { replacementOps } from './replacement-ops.ts';
import { writeRejectionReason } from './command-protection.ts';

interface Edit {
  start: number;
  end: number;
  text: string;
  owner: string;
}
const unsupported = (reason: string) => ({ ok: false, code: 'unsupported', reason }) as const;

/** One local selection, one guarded transaction. Run boundaries and inline objects survive. */
export function changeSelectionCase(
  surface: PaginatedSurface,
  read: (id: string) => ParagraphText | null,
  mode: CaseMode,
  commit: boolean
): ExecResult | null {
  if (commit) surface.flushPendingInput();
  const {
    selection: { anchor, head },
    cellSelection,
  } = surface.state();
  if (cellSelection || (anchor.paragraphId === head.paragraphId && anchor.offset === head.offset))
    return unsupported('select text to change its case');
  const scope = surface.storyScope();
  const order =
    anchor.paragraphId === head.paragraphId
      ? [head.paragraphId]
      : surface.activeScope().kind === 'frame'
        ? [...openTextboxParagraphOrder(surface).keys()]
        : surface.session.paragraphIdsIn(scope);
  const a = order.indexOf(anchor.paragraphId);
  const b = order.indexOf(head.paragraphId);
  if (a < 0 || b < 0) return unsupported('select text in a single story');
  const forward = a < b || (a === b && anchor.offset <= head.offset);
  const [first, last] = forward ? [anchor, head] : [head, anchor];
  const ops: TreeDocOp[] = [];
  let lastLength = 0;
  for (const id of order.slice(Math.min(a, b), Math.max(a, b) + 1)) {
    const part = partOfNodeId(surface.session, id);
    const node = part && findNode(part, id);
    const paragraph = read(id);
    if (!part || node?.kind !== 'paragraph' || !paragraph)
      return unsupported('paragraph is unavailable');
    if (id === last.paragraphId) lastLength = paragraph.text.length;
    const start = id === first.paragraphId ? acceptedOffsetFromText(paragraph, first.offset) : 0;
    const end =
      id === last.paragraphId
        ? acceptedOffsetFromText(paragraph, last.offset)
        : paragraph.acceptedText.length;
    if (end <= start) continue;
    const letters = caseLetters(paragraph.acceptedText, start, end, mode, paragraph.proofing);
    if (!letters)
      return unsupported('the selection splits a character or has an unsupported language tag');
    const spans = paragraph.acceptedSpans;
    const segments = segmentsOf(node);
    const edits: Edit[] = [];
    let si = 0,
      ri = 0;
    for (const letter of letters) {
      while (spans[si] && spans[si]!.acceptedStart + spans[si]!.length <= letter.start) si++;
      const span = spans[si];
      if (!span || letter.end > span.acceptedStart + span.length)
        return unsupported('a changed character crosses an excluded text boundary');
      const from = span.start + letter.start - span.acceptedStart;
      const to = from + letter.end - letter.start;
      while (segments[ri] && segments[ri]!.end <= from) ri++;
      const leaf = segments[ri];
      if (
        !leaf ||
        from < leaf.start ||
        to > leaf.end ||
        leaf.node.kind !== 'textValue' ||
        leaf.node.value.slice(from - leaf.start, to - leaf.start) !== paragraph.text.slice(from, to)
      )
        return unsupported('a changed character crosses a native text boundary');
      const previous = edits.at(-1);
      if (previous && previous.owner === leaf.node.id && previous.end === from) {
        previous.end = to;
        previous.text += letter.text;
      } else edits.push({ start: from, end: to, text: letter.text, owner: leaf.node.id });
    }
    for (const edit of edits.reverse()) {
      const landing = surface.replacementLanding(id, edit.start, edit.end);
      const refusal = matchRangeSkip(part, id, edit.start, edit.end, landing ?? edit.start);
      if (refusal)
        return {
          ok: false,
          code: refusal === 'bound' || refusal === 'locked' ? refusal : 'unsupported',
          reason: writeRejectionReason(refusal, surface.session.settingsRoot()),
        };
      ops.push(...replacementOps(id, edit.start, edit.end, edit.text, landing));
    }
  }
  if (!ops.length || !commit) return null;
  const result = surface.applyAutomationOps(
    () => ops,
    scope,
    undefined,
    false,
    (selection) => {
      const after = read(last.paragraphId);
      if (!after) return selection;
      const end = { ...last, offset: last.offset + after.text.length - lastLength };
      return forward ? { anchor: first, head: end } : { anchor: end, head: first };
    }
  );
  if (!result.rejected) return null;
  return {
    ok: false,
    code: 'unsupported',
    reason: surface.state().lastRejection ?? String(result.reason),
  };
}
