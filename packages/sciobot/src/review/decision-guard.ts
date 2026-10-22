import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { NativeEditFailure } from '../contracts/document';

const inFlight = new WeakSet<DocxEditorInstance>();

/** Application review decisions share one per-editor reentry guard. */
export function withReviewDecisionGuard<Result>(
  editor: DocxEditorInstance,
  run: () => Result
): Result | NativeEditFailure {
  if (inFlight.has(editor))
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'A review group decision is already running.',
    };
  inFlight.add(editor);
  try {
    return run();
  } finally {
    inFlight.delete(editor);
  }
}
