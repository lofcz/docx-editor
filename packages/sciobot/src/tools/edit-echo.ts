import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { createEditPlan } from '../diff/plan';
import type { MarkdownProjection } from '../projection/markdown';

export type EditEchoProjection = Pick<MarkdownProjection, 'text' | 'revision' | 'displayMode'>;
export interface EditEcho {
  readonly revision: RevisionToken;
  /** Exact landed Markdown. No line numbers or truncation markers enter this copyable slice. */
  readonly excerpt: string;
  readonly startLine: number;
  readonly endLine: number;
  readonly totalLines: number;
  readonly truncated: boolean;
  /** Application tool feedback, including its document_excerpt wrapper. */
  readonly message: string;
}
export type EditEchoResult =
  | { readonly success: true; readonly echo: EditEcho }
  | NativeEditFailure;
export const EDIT_ECHO_CONTEXT_LINES = 2;
export const EDIT_ECHO_MAX_LINES = 40;

function snapshot(value: EditEchoProjection): EditEchoProjection | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const fields = Object.getOwnPropertyDescriptors(value);
  for (const key of ['text', 'revision', 'displayMode'])
    if (!fields[key]?.enumerable || !('value' in fields[key]!)) return null;
  const text = fields.text!.value,
    displayMode = fields.displayMode!.value;
  const rawRevision = fields.revision!.value;
  if (!rawRevision) return null;
  const token = createEditPlan(rawRevision, []);
  if (typeof text !== 'string' || !token.success || typeof displayMode !== 'string') return null;
  return {
    text,
    revision: token.plan.revision,
    displayMode: displayMode as MarkdownProjection['displayMode'],
  };
}

/**
 * Compare a pre-edit working projection with a fresh, landed working projection.
 * The caller reads current after flushing native input. Never pass the requested target as landed.
 * Existing native, block, or opaque MDX projectors must use the same options on both snapshots.
 * This formatter performs no document reads or writes; I03-I05 own the live bridge.
 */
export function buildEditEcho(
  before: EditEchoProjection,
  landed: EditEchoProjection,
  current: RevisionToken
): EditEchoResult {
  const fail = (
    error: string,
    code: NativeEditFailure['code'] = 'invalid-input'
  ): NativeEditFailure => ({ success: false, code, error });
  let old: EditEchoProjection | null, after: EditEchoProjection | null;
  const token = createEditPlan(current, []);
  try {
    old = snapshot(before);
    after = snapshot(landed);
  } catch {
    return fail('Invalid echo projection.');
  }
  if (!old || !after || !token.success) return fail('Invalid echo projection or revision.');
  const actual = token.plan.revision;
  if (
    !sameRevision(after.revision, actual) ||
    old.revision.documentId !== actual.documentId ||
    old.revision.epoch !== actual.epoch ||
    old.revision.revision > actual.revision
  )
    return Object.freeze({
      success: false,
      code: 'stale-revision',
      error: 'Read the landed working projection before building the edit echo.',
      expectedRevision: Object.freeze({ ...after.revision }),
      actualRevision: actual,
    });
  if (old.displayMode !== 'proposed' || after.displayMode !== 'proposed')
    return fail('Edit echoes require working projections.', 'unsupported-operation');
  if (old.revision.revision === after.revision.revision && old.text !== after.text)
    return fail('One revision has inconsistent echo projections.', 'invalid-document');
  // Application-owned excerpt rules. See fixtures/edit-echo.json for source-executed expectations.
  const beforeLines = old.text.split('\n'),
    afterLines = after.text.split('\n');
  let head = 0;
  while (
    head < beforeLines.length &&
    head < afterLines.length &&
    beforeLines[head] === afterLines[head]
  )
    head++;
  let tail = 0;
  while (
    tail < beforeLines.length - head &&
    tail < afterLines.length - head &&
    beforeLines[beforeLines.length - 1 - tail] === afterLines[afterLines.length - 1 - tail]
  )
    tail++;
  const firstChanged = head + 1;
  const lastChanged = Math.max(afterLines.length - tail, firstChanged);
  const startLine = Math.max(1, firstChanged - EDIT_ECHO_CONTEXT_LINES);
  let endLine = Math.min(afterLines.length, lastChanged + EDIT_ECHO_CONTEXT_LINES);
  const truncated = endLine - startLine + 1 > EDIT_ECHO_MAX_LINES;
  if (truncated) endLine = startLine + EDIT_ECHO_MAX_LINES - 1;
  const excerpt = afterLines.slice(startLine - 1, endLine).join('\n');
  const suffix = truncated ? '\n… (excerpt truncated — use read_document for the rest)' : '';
  return Object.freeze({
    success: true,
    echo: Object.freeze({
      revision: Object.freeze({ ...after.revision }),
      excerpt,
      startLine,
      endLine,
      totalLines: afterLines.length,
      truncated,
      message: `Updated document excerpt around your change (plain text — you may copy old_string for a follow-up edit directly from it):\n<document_excerpt>\n${excerpt}${suffix}\n</document_excerpt>`,
    }),
  });
}
