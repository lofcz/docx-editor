import type { HeadlessDocumentView } from '@docx-editor.dev/core/store';
import type { RevisionToken } from '../contracts/document';
import {
  readNativeReviewView,
  type ReviewViewOptions,
  type ReviewViewResult,
} from './document-view';

/** Read native content and formatting as if pending revisions were rejected, without decisions. */
export function readCommittedView(
  source: HeadlessDocumentView,
  revision: RevisionToken,
  options?: ReviewViewOptions
): Promise<ReviewViewResult> {
  return readNativeReviewView(source, revision, 'committed', options);
}
