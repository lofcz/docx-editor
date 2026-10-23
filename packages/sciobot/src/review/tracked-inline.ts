import { withReviewGroupRevisionReservations } from './group-allocation';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { segmentGraphemes } from '@docx-editor.dev/core/layout';
import {
  findNode,
  paragraphTextOf,
  trackedReplacementLanding,
  type OoxmlParagraphNode,
  type OoxmlPart,
  type RevisionAddress,
  type ReviewRevisionItem,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import { PARAGRAPH_TEXT_LIMITS, type PlannedTextEdit } from '../diff/paragraph-text';
import { serializeEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import {
  checkNativeCellScope,
  nearestCellAncestor,
  type NativeCellScope,
} from '../diff/cell-scope';
import { trackedRunFormatOps } from './tracked-run-format';
import { leadingReplacementEdit } from './leading-replacements';
import { persistReviewGroups } from './groups';

export interface TrackTextContext {
  /** The current mount supplies these values independently of the plan. */
  readonly documentId: string;
  readonly epoch: string;
  readonly author: string;
  /** One UTC ISO-8601 timestamp for this proposal batch. */
  readonly date: string;
}

/** Native addresses survive save/reopen; live queue keys remain editor-instance data. */
export interface NativeTextProposal {
  readonly identity: string;
  readonly partName: string;
  readonly kind: ReviewRevisionItem['revisionKind'];
  readonly addresses: readonly RevisionAddress[];
  readonly text: string;
  readonly replacedText: string;
  readonly formattingKind?: string;
  readonly formattingChanges?: ReviewRevisionItem['formattingChanges'];
}

export type TrackTextResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposals: readonly NativeTextProposal[];
      readonly receiptError?: string;
    }
  | NativeEditFailure;

const loneSurrogate = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const invalidXml = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;
const invalidInserted = /[\u0000-\u0008\u000B-\u001D\uFFFE\uFFFF]/;

export function canonicalTrackedDate(value: string): string | null {
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(value)) return null;
  const timestamp = Date.parse(value);
  if (!Number.isFinite(timestamp)) return null;
  const normalized = value.includes('.')
    ? value.replace(/\.(\d{1,3})Z$/, (_, fraction: string) => `.${fraction.padEnd(3, '0')}Z`)
    : value.slice(0, -1) + '.000Z';
  const date = new Date(timestamp).toISOString();
  return date === normalized ? date : null;
}

function proposal(item: ReviewRevisionItem, fallbackPart: string): NativeTextProposal {
  const partName = item.ranges[0]?.partName ?? fallbackPart;
  const addresses = item.addresses.map((address) => Object.freeze({ ...address }));
  const keys = addresses
    .map((address) => JSON.stringify([address.id, address.author, address.date ?? null]))
    .sort();
  return Object.freeze({
    identity: JSON.stringify([
      partName,
      item.revisionKind,
      keys,
      ...(item.formattingKind ? [item.formattingKind] : []),
    ]),
    partName,
    kind: item.revisionKind,
    addresses: Object.freeze(addresses),
    text: item.text,
    replacedText: item.replacedText,
    ...(item.formattingKind ? { formattingKind: item.formattingKind } : {}),
    ...(item.formattingChanges
      ? {
          formattingChanges: Object.freeze(
            item.formattingChanges.map((change) => Object.freeze({ ...change }))
          ),
        }
      : {}),
  });
}

export function trackedParagraphTextOps(
  edit: PlannedTextEdit,
  paragraph: OoxmlParagraphNode,
  author: string,
  date: string
): TreeDocOp[] {
  const revision = { author, date };
  let offset = trackedReplacementLanding(paragraph, edit.start, edit.end, author, date);
  const ops: TreeDocOp[] = [];
  if (edit.end > edit.start)
    ops.push({
      op: 'deleteText',
      paragraphId: edit.paragraphId,
      start: edit.start,
      end: edit.end,
      revision,
    });
  for (const text of edit.after.split(/(\n|\t)/)) {
    if (!text) continue;
    if (text === '\n')
      ops.push({ op: 'insertHardBreak', paragraphId: edit.paragraphId, offset, revision });
    else if (text === '\t')
      ops.push({ op: 'insertTab', paragraphId: edit.paragraphId, offset, revision });
    else
      ops.push({
        op: 'insertText',
        paragraphId: edit.paragraphId,
        offset,
        text,
        revision,
        bias: edit.before ? 'left' : (edit.bias ?? 'left'),
      });
    offset += text.length;
  }
  return ops;
}

/**
 * Shared guarded body transaction. A native cell scope also enables run-format plans.
 * Mount identity and the current package revision must match. The installed editor needs
 * SCIOBOT_NATIVE_REVIEW_MODULE (or an equivalent review contribution). No editing mode changes occur.
 * Broader projection protection and streaming/session reconciliation remain separate layers.
 */
export function trackInlinePlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext,
  cellScope?: NativeCellScope
): TrackTextResult {
  return withReviewGroupRevisionReservations(editor.surface?.session.currentPackage() ?? null, () =>
    trackInlinePlanReserved(editor, input, context, cellScope)
  );
}

function trackInlinePlanReserved(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext,
  cellScope?: NativeCellScope
): TrackTextResult {
  const fail = (
    error: string,
    code: NativeEditFailure['code'] = 'invalid-input'
  ): NativeEditFailure => ({ success: false, code, error });
  const contextProblem = jsonDataProblem(context, { maxDepth: 4, maxNodes: 20 });
  if (contextProblem) return fail(contextProblem);
  if (
    !context ||
    typeof context.documentId !== 'string' ||
    !context.documentId ||
    typeof context.epoch !== 'string' ||
    !context.epoch ||
    typeof context.author !== 'string' ||
    !context.author.trim() ||
    context.author.length > 4096 ||
    invalidXml.test(context.author) ||
    loneSurrogate.test(context.author) ||
    typeof context.date !== 'string'
  )
    return fail('Invalid tracked text context.');
  const date = canonicalTrackedDate(context.date);
  if (!date) return fail('Tracked text needs a valid UTC ISO-8601 date.');
  let plan: CanonicalEditPlan;
  try {
    plan = JSON.parse(serializeEditPlan(input)) as CanonicalEditPlan;
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause));
  }
  const actual = readEditorRevision(editor, context.documentId, context.epoch);
  const surface = editor.surface;
  if (!actual || !surface) return fail('The native editor is not attached.', 'document-not-ready');
  if (!sameRevision(plan.revision, actual))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read the current document before proposing this edit.',
      expectedRevision: plan.revision,
      actualRevision: actual,
    };
  if (surface.editingMode() === 'view')
    return fail('The document is open for viewing.', 'unsupported-operation');
  const session = surface.session,
    part = session.part(),
    beforePackage = session.currentPackage();
  const prepared = prepareTrackedInlinePlan(
    part,
    plan,
    context,
    date,
    cellScope,
    !session.trackingSettings().doNotTrackFormatting &&
      editor.snapshot().revisionMarkup.trackFormatting
  );
  if (!prepared.success) return prepared;
  const { ops, paragraphIds } = prepared;
  if (!ops.length)
    return { success: true, revision: Object.freeze(actual), proposals: Object.freeze([]) };
  let before: Map<string, string>;
  const receipts = () =>
    session
      .reviewItems()
      .filter(
        (item): item is ReviewRevisionItem =>
          item.kind === 'revision' &&
          item.author === context.author &&
          ['insert', 'delete', 'replace', ...(cellScope !== undefined ? ['format'] : [])].includes(
            item.revisionKind
          ) &&
          item.ranges.some(
            (range) => range.partName === part.name && paragraphIds.has(range.start.paragraphId)
          )
      )
      .map((item) => proposal(item, part.name));
  try {
    before = new Map(
      receipts().map((item) => [
        item.identity,
        JSON.stringify([item.text, item.replacedText, item.formattingChanges ?? null]),
      ])
    );
  } catch (cause) {
    return fail(
      `Unable to read native review state: ${cause instanceof Error ? cause.message : String(cause)}`,
      'invalid-document'
    );
  }
  try {
    const result = surface.applyAutomationOps(
      () => ops,
      { kind: 'body' },
      [
        leadingReplacementEdit(part, paragraphIds, context.author, date),
        (pkg) =>
          persistReviewGroups(
            beforePackage,
            pkg,
            [],
            editor.snapshot().collaborationStatus === 'inactive'
          ),
      ],
      true,
      true, // Each proposal needs its own unit, including within the collaboration capture window.
      'after'
    );
    if (result.rejected || !result.committed)
      return fail(
        String(result.reason ?? 'The tracked edit did not commit.'),
        result.reason === 'review-module-required' ? 'unsupported-operation' : 'apply-failed'
      );
  } catch (cause) {
    const changed = session.packageRevision() !== actual.revision;
    return {
      success: false,
      code: 'apply-failed',
      error: `${changed ? 'The document changed during a failed native call. Re-read before retrying. ' : ''}${cause instanceof Error ? cause.message : String(cause)}`,
      expectedRevision: actual,
      actualRevision: { ...actual, revision: session.packageRevision() },
    };
  }
  const revision = Object.freeze({ ...actual, revision: session.packageRevision() });
  try {
    const proposals = receipts().filter(
      (item) =>
        before.get(item.identity) !==
        JSON.stringify([item.text, item.replacedText, item.formattingChanges ?? null])
    );
    return { success: true, revision, proposals: Object.freeze(proposals) };
  } catch (cause) {
    return {
      success: true,
      revision,
      proposals: Object.freeze([]),
      receiptError: `The edit committed, but its review receipt is unavailable: ${cause instanceof Error ? cause.message : String(cause)}`,
    };
  }
}

/** Pure native lowering shared with dependent owned-proposal replacement stages. */
export function prepareTrackedInlinePlan(
  part: OoxmlPart,
  plan: CanonicalEditPlan,
  context: Pick<TrackTextContext, 'author'>,
  date: string,
  cellScope?: NativeCellScope,
  trackFormatting = true
):
  | {
      readonly success: true;
      readonly ops: TreeDocOp[];
      readonly paragraphIds: ReadonlySet<string>;
    }
  | NativeEditFailure {
  const fail = (
    error: string,
    code: NativeEditFailure['code'] = 'invalid-input'
  ): NativeEditFailure => ({ success: false, code, error });
  if (cellScope !== undefined) {
    const problem = checkNativeCellScope(part, cellScope);
    if (problem) return problem;
  }
  const groups = new Map<string, PlannedTextEdit[]>();
  const formatEdits: Extract<CanonicalEditPlan['operations'][number], { kind: 'format' }>[] = [];
  const paragraphIds = new Set<string>();
  let atoms = 0;
  for (const edit of plan.operations) {
    if (edit.kind === 'text' && edit.insideHyperlink !== undefined)
      return fail(
        'Tracked hyperlink-owned insertion needs native revision ownership.',
        'unsupported-operation'
      );
    if (
      edit.partName !== part.name ||
      (edit.kind !== 'text' &&
        !(cellScope !== undefined && edit.kind === 'format' && edit.target.kind === 'run'))
    )
      return fail('This executor accepts body paragraph text plans only.', 'unsupported-operation');
    const id =
      edit.kind === 'text'
        ? edit.paragraphId
        : edit.target.kind === 'run'
          ? edit.target.paragraphId
          : '';
    if (
      cellScope !== undefined &&
      nearestCellAncestor(part, id, 'tableCell')?.id !== cellScope.cellId
    )
      return fail(
        'A planned paragraph does not belong to the requested native cell.',
        'ambiguous-target'
      );
    paragraphIds.add(id);
    if (edit.kind === 'format') {
      formatEdits.push(edit);
      continue;
    }
    if (
      loneSurrogate.test(edit.after) ||
      invalidInserted.test(edit.after) ||
      edit.after.length > PARAGRAPH_TEXT_LIMITS.maxBytes ||
      new TextEncoder().encode(edit.after).byteLength > PARAGRAPH_TEXT_LIMITS.maxBytes
    )
      return fail('Invalid inserted paragraph text.');
    for (const character of edit.after)
      if (
        (character === '\t' ||
          character === '\n' ||
          character === '\u001e' ||
          character === '\u001f') &&
        ++atoms > PARAGRAPH_TEXT_LIMITS.maxInlineAtoms
      )
        return fail('Inserted text exceeds its inline element limit.', 'unsupported-operation');
    const group = groups.get(edit.paragraphId) ?? [];
    group.push(edit);
    groups.set(edit.paragraphId, group);
  }
  if (formatEdits.length && !trackFormatting)
    return fail(
      'Formatting tracking is disabled for this document or editor.',
      'unsupported-operation'
    );
  const formatted = trackedRunFormatOps(part, formatEdits, { author: context.author, date });
  if (!formatted.success) return formatted;
  const ops: TreeDocOp[] = [...formatted.ops];
  for (const [id, edits] of groups) {
    const paragraph = findNode(part, id),
      text = paragraphTextOf(part, id);
    if (paragraph?.kind !== 'paragraph' || text === null)
      return fail('A planned paragraph no longer exists.', 'invalid-document');
    const boundaries = new Set([
      0,
      text.length,
      ...segmentGraphemes(text).flatMap((segment) => [segment.utf16From, segment.utf16To]),
    ]);
    for (const edit of edits) {
      if (edit.end > text.length || text.slice(edit.start, edit.end) !== edit.before)
        return fail('Planned before-text does not match the current paragraph.', 'stale-revision');
      if (!boundaries.has(edit.start) || !boundaries.has(edit.end))
        return fail('A tracked edit splits a native grapheme boundary.', 'ambiguous-target');
    }
    const ordered = [...edits].sort((a, b) => b.start - a.start);
    const merged: PlannedTextEdit[] = [];
    for (const edit of ordered) {
      const right = merged.at(-1);
      if (right && (edit.end > right.start || edit.start === right.start))
        return fail('Planned text ranges overlap.', 'ambiguous-target');
      if (right && edit.end === right.start)
        merged[merged.length - 1] = {
          ...edit,
          end: right.end,
          before: edit.before + right.before,
          after: edit.after + right.after,
        };
      else merged.push(edit);
    }
    for (const edit of merged)
      ops.push(...trackedParagraphTextOps(edit, paragraph, context.author, date));
  }
  return { success: true, ops, paragraphIds };
}
