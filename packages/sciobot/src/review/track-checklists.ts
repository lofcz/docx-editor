import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { CanonicalEditPlan } from '../diff/plan';
import { fail } from './tracked-structure';
import type { TrackTextContext } from './tracked-inline';
import { replaceOwnedProposals } from './replace-owned-proposals';
/** Propose checklist changes through the same atomic native review coordinator. */
export function trackChecklistPlan(
  editor: DocxEditorInstance,
  plan: CanonicalEditPlan,
  context: TrackTextContext
) {
  if (plan.operations.some((edit) => edit.kind !== 'list'))
    return fail('This executor accepts checklist membership plans only.', 'unsupported-operation');
  return replaceOwnedProposals(editor, [], plan, context);
}
