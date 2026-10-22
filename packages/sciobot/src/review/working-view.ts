import type { HeadlessDocumentView } from '@docx-editor.dev/core/store';
import type { RevisionToken } from '../contracts/document';
import {
  readNativeReviewView,
  type ReviewViewOptions,
  type ReviewViewResult,
} from './document-view';

/** Read the proposed native result while preserving every pending revision in the live store. */
export function readWorkingView(
  source: HeadlessDocumentView,
  revision: RevisionToken,
  options?: ReviewViewOptions
): Promise<ReviewViewResult> {
  return readNativeReviewView(source, revision, 'working', options);
}
