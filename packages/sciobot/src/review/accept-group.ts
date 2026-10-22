import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { RevisionToken } from '../contracts/document';
import { decideReviewGroup, type GroupDecisionResult } from './decide-group';
export type { GroupDecisionResult } from './decide-group';

/** Accept the complete current group in one native review transaction and one undo unit. */
export function acceptReviewGroup(
  editor: DocxEditorInstance,
  groupId: string,
  expectedRevision: RevisionToken,
  identity: Pick<RevisionToken, 'documentId' | 'epoch'>
): GroupDecisionResult {
  return decideReviewGroup(editor, groupId, expectedRevision, identity, 'accept');
}
