import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { NativeCellScope } from '../diff/cell-scope';
import type { CanonicalEditPlan } from '../diff/plan';
import { trackInlinePlan, type TrackTextContext, type TrackTextResult } from './tracked-inline';

export type { NativeCellScope } from '../diff/cell-scope';

/** Apply text and run-mark plans inside one live native cell, with one review/history transaction. */
export function trackTableCellPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  scope: NativeCellScope,
  context: TrackTextContext
): TrackTextResult {
  if (!scope)
    return {
      success: false,
      code: 'invalid-input',
      error: 'A tracked cell edit needs its native cell scope.',
    };
  return trackInlinePlan(editor, input, context, scope);
}
