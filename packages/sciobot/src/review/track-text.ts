import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { CanonicalEditPlan } from '../diff/plan';
import { trackInlinePlan, type TrackTextContext, type TrackTextResult } from './tracked-inline';

export type { TrackTextContext, TrackTextResult, NativeTextProposal } from './tracked-inline';

/** Apply body paragraph text plans as native revisions through one guarded transaction. */
export function trackTextPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext
): TrackTextResult {
  return trackInlinePlan(editor, input, context);
}
