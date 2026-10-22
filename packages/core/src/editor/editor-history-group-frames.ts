// Native writes that extend one explicit history group.
//
// `Editor.exec` binds a handle to one command. An application adapter can need two native
// writes for one gesture: a surface command, then a repair that only the adapter knows how to
// build. Each `run` binds the handle for one write, so both land in the handle's undo step.
import type { HistoryGroup, HistoryGroupOutcome } from '../contracts/editor.ts';
import type { DocxEditorInstance } from './docx-editor-types.ts';
import { editorAutomationFrames } from './editor-history-group-registry.ts';

/**
 * Frames of one open history group.
 *
 * @internal
 */
export interface HistoryGroupFrames {
  /** Bind the group for the native writes inside `write`. */
  run<T>(write: () => T): T;
  /** Settle one outcome on the handle. Safe to call more than once. */
  finish(): HistoryGroupOutcome | null;
}

/**
 * Start frames for `group`, or answer `null` for a closed handle or one from another editor.
 * Surface commands and `applyAutomationOps` writes inside `run` extend the same undo step.
 *
 * @internal
 */
export function beginHistoryGroupFrames(
  editor: DocxEditorInstance,
  group: HistoryGroup
): HistoryGroupFrames | null {
  const surface = editor.surface;
  const frames = surface ? editorAutomationFrames(editor, group) : null;
  if (!surface || !frames) return null;
  let finished = false;
  let outcome: HistoryGroupOutcome | null = null;
  return {
    run: (write) => frames.bind(surface, write),
    finish: () => {
      if (!finished) {
        finished = true;
        outcome = frames.finish();
      }
      return outcome;
    },
  };
}
