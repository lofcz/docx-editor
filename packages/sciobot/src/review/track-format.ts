import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { serializeEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { fail } from './tracked-structure';
import type { TrackTextContext } from './tracked-inline';
import { replaceOwnedProposals } from './replace-owned-proposals';

/** Propose native run, paragraph, and list properties through one review transaction. */
export function trackFormatPlan(
  editor: DocxEditorInstance,
  plan: CanonicalEditPlan,
  context: TrackTextContext
) {
  let copied: CanonicalEditPlan;
  try {
    copied = JSON.parse(serializeEditPlan(plan)) as CanonicalEditPlan;
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'invalid-input');
  }
  if (copied.operations.some((edit) => edit.kind !== 'format' && edit.kind !== 'list'))
    return fail('This executor accepts formatting and list plans only.', 'unsupported-operation');
  return replaceOwnedProposals(editor, [], copied, context);
}
