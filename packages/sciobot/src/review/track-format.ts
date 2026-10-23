import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { CanonicalEditPlan } from '../diff/plan';
import { fail } from './tracked-structure';
import type { TrackTextContext } from './tracked-inline';
import { replaceOwnedProposals } from './replace-owned-proposals';

/** Propose native run, paragraph, and list properties through one review transaction. */
export function trackFormatPlan(
  editor: DocxEditorInstance,
  plan: CanonicalEditPlan,
  context: TrackTextContext
) {
  if (plan.operations.some((edit) => edit.kind !== 'format' && edit.kind !== 'list'))
    return fail('This executor accepts formatting and list plans only.', 'unsupported-operation');
  return replaceOwnedProposals(editor, [], plan, context);
}
