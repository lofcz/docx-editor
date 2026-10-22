import type { TreeDocOp } from '@docx-editor.dev/core/store';
import { segmentGraphemes } from '@docx-editor.dev/core/layout';
import { diffArrays } from 'diff';
import type { RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';

export interface ParagraphTextSource {
  readonly revision: RevisionToken;
  readonly partName: string;
  readonly paragraphId: string;
  /** Canonical model text, or one exactly mapped contiguous model fragment. LF represents a manual line break. */
  readonly text: string;
  /** Model offset of a fragment. Omit for a complete paragraph. */
  readonly offset?: number;
}

export type PlannedTextEdit = Extract<CanonicalEdit, { kind: 'text' }>;
export const PARAGRAPH_TEXT_LIMITS = Object.freeze({
  maxBytes: 2 * 1024 * 1024,
  diffTimeoutMs: 1000,
  maxInlineAtoms: 4096,
});
const tokens = (text: string): string[] => segmentGraphemes(text).map((entry) => entry.text);
const loneSurrogate = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
// Model hyphens U+001E/U+001F become native hyphen elements. LF and tab have explicit native operations.
const invalidInsertedCharacter = /[\u0000-\u0008\u000B-\u001D\uFFFE\uFFFF]/;
const wordCharacter = /[\p{L}\p{N}\p{M}\p{Extended_Pictographic}]/u;

/**
 * Diff paragraph model text without rebuilding its runs. Edits use original UTF-16 coordinates, last first.
 * Review callers resolve visible fragments to contiguous model coordinates before invoking this planner.
 * Sibling plans own paragraph creation; LF inside this input means a manual line break.
 */
export function planParagraphText(source: ParagraphTextSource, target: string): EditPlanResult {
  const fail = (
    error: string,
    code: 'invalid-input' | 'unsupported-operation' = 'invalid-input'
  ): EditPlanResult => ({ success: false, code, error });
  const problem = jsonDataProblem({ source, target }, { maxDepth: 8, maxNodes: 100 });
  if (problem) return fail(problem);
  if (
    !source ||
    typeof source.text !== 'string' ||
    typeof target !== 'string' ||
    typeof source.partName !== 'string' ||
    !source.partName.startsWith('/') ||
    source.partName.split('/').includes('..') ||
    typeof source.paragraphId !== 'string' ||
    !source.paragraphId ||
    (source.offset !== undefined && (!Number.isSafeInteger(source.offset) || source.offset < 0))
  )
    return fail('Invalid paragraph text source.');
  const offset = source.offset ?? 0;
  if (!Number.isSafeInteger(offset + source.text.length + target.length))
    return fail('Paragraph model offsets exceed the safe integer range.');
  if (loneSurrogate.test(source.text) || loneSurrogate.test(target))
    return fail('Paragraph text contains an unpaired surrogate.');
  for (const text of [source.text, target])
    if (
      text.length > PARAGRAPH_TEXT_LIMITS.maxBytes ||
      new TextEncoder().encode(text).byteLength > PARAGRAPH_TEXT_LIMITS.maxBytes
    )
      return fail('Paragraph text exceeds its byte limit.');
  if (source.text === target) return createEditPlan(source.revision, []);

  const changes = diffArrays(tokens(source.text), tokens(target), {
    timeout: PARAGRAPH_TEXT_LIMITS.diffTimeoutMs,
  });
  if (!changes)
    return fail('Paragraph text diff exceeded its planning budget.', 'unsupported-operation');
  const edits: PlannedTextEdit[] = [];
  let cursor = offset,
    start = offset,
    before = '',
    after = '';
  const flush = () => {
    if (before || after) {
      const local = start - offset;
      const left = local
        ? String.fromCodePoint(
            source.text.codePointAt(
              local -
                (source.text.charCodeAt(local - 1) >= 0xdc00 &&
                source.text.charCodeAt(local - 1) <= 0xdfff
                  ? 2
                  : 1)
            )!
          )
        : '';
      const right =
        local < source.text.length ? String.fromCodePoint(source.text.codePointAt(local)!) : '';
      // A word prefix joins its following run; a word suffix joins its preceding run.
      const bias =
        before || (wordCharacter.test(right) && !wordCharacter.test(left)) ? 'right' : 'left';
      edits.push({
        kind: 'text',
        partName: source.partName,
        paragraphId: source.paragraphId,
        start,
        end: cursor,
        before,
        after,
        bias,
      });
    }
    before = after = '';
  };
  for (const change of changes) {
    const text = change.value.join('');
    if (!change.added && !change.removed) {
      flush();
      cursor += text.length;
      start = cursor;
      continue;
    }
    if (change.removed) {
      before += text;
      cursor += text.length;
    } else after += text;
  }
  flush();
  if (edits.some((edit) => loneSurrogate.test(edit.before) || loneSurrogate.test(edit.after)))
    return fail(
      'The native grapheme strategy produced a split surrogate boundary.',
      'unsupported-operation'
    );
  if (edits.some((edit) => invalidInsertedCharacter.test(edit.after)))
    return fail('Inserted paragraph text contains an invalid model character.');
  let atoms = 0;
  for (const edit of edits)
    for (const character of edit.after)
      if (
        (character === '\t' ||
          character === '\n' ||
          character === '\u001e' ||
          character === '\u001f') &&
        ++atoms > PARAGRAPH_TEXT_LIMITS.maxInlineAtoms
      )
        return fail(
          'Inserted paragraph text exceeds its inline element limit.',
          'unsupported-operation'
        );
  // Original coordinates remain valid when a later edit commits before an earlier edit.
  edits.reverse();
  return createEditPlan(source.revision, edits);
}

/**
 * Lower one validated semantic text edit for direct application. This function does not write.
 * Insert before removing old text so a replacement keeps the original run's formatting and identity.
 * The executor must check revision, current before-text, protection, and the part before applying these operations.
 */
export function directParagraphTextOps(edit: PlannedTextEdit): readonly TreeDocOp[] {
  const ops: TreeDocOp[] = [];
  let offset = edit.start;
  for (const text of edit.after.split(/(\n|\t)/)) {
    if (!text) continue;
    if (text === '\n') ops.push({ op: 'insertHardBreak', paragraphId: edit.paragraphId, offset });
    else if (text === '\t') ops.push({ op: 'insertTab', paragraphId: edit.paragraphId, offset });
    else
      ops.push({
        op: 'insertText',
        paragraphId: edit.paragraphId,
        offset,
        text,
        bias: edit.bias ?? (edit.before ? 'right' : 'left'),
      });
    offset += text.length;
  }
  if (edit.end > edit.start)
    ops.push({
      op: 'deleteText',
      paragraphId: edit.paragraphId,
      start: edit.start + edit.after.length,
      end: edit.end + edit.after.length,
    });
  return Object.freeze(ops.map((op) => Object.freeze(op)));
}
