import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { RevisionToken } from '../contracts/document';
import { decideReviewGroup, type GroupDecisionResult } from './decide-group';

/** Reject every current group against live content, with a fresh revision and one undo unit. */
export function rejectAllReviewGroups(
  editor: DocxEditorInstance,
  expectedRevision: RevisionToken,
  identity: Pick<RevisionToken, 'documentId' | 'epoch'>
): GroupDecisionResult {
  return decideReviewGroup(editor, null, expectedRevision, identity, 'reject');
}
