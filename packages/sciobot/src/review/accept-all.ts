import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { RevisionToken } from '../contracts/document';
import { decideReviewGroup, type GroupDecisionResult } from './decide-group';

/** Accept every complete current group atomically, with a fresh revision and one undo unit. */
export function acceptAllReviewGroups(
  editor: DocxEditorInstance,
  expectedRevision: RevisionToken,
  identity: Pick<RevisionToken, 'documentId' | 'epoch'>
): GroupDecisionResult {
  return decideReviewGroup(editor, null, expectedRevision, identity, 'accept');
}
