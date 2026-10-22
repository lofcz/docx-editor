import type { EditorCollaborationSession } from '@docx-editor.dev/core/collaboration';

// Both private entry bundles use this runtime capability. A module-local WeakMap would split it.
export const REVIEW_HISTORY_CAPABILITY = Symbol.for('sciobot.docx.review-history.4');

export interface ReviewHistoryGuard {
  (): void;
  capture<T>(run: () => T, flagWrites: number, fieldWrites?: number): T;
}

/** Keep native removal flags and their ownership intents in the same native history frame. */
export function captureReviewGroupHistory<T>(
  session: EditorCollaborationSession | null,
  run: () => T,
  flagWrites: number,
  fieldWrites = 0
): T {
  if (!session) return run();
  const guard = (session as unknown as Record<symbol, unknown>)[REVIEW_HISTORY_CAPABILITY];
  if (typeof guard !== 'function' || typeof (guard as ReviewHistoryGuard).capture !== 'function')
    throw new Error('Shared group decisions require the complete owned review history scope.');
  return (guard as ReviewHistoryGuard).capture(run, flagWrites, fieldWrites);
}

/** Require both text-attribute ownership and removal-flag ownership before shared decisions. */
export function reviewGroupHistoryRefusal(
  session: EditorCollaborationSession | null
): string | null {
  if (!session) return null;
  try {
    const guard = (session as unknown as Record<symbol, unknown>)[REVIEW_HISTORY_CAPABILITY];
    if (typeof guard !== 'function' || typeof (guard as ReviewHistoryGuard).capture !== 'function')
      return 'Shared group decisions require an owned review history scope attached before room bootstrap.';
    guard();
    return null;
  } catch (cause) {
    return cause instanceof Error ? cause.message : String(cause);
  }
}
