import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { CanonicalEditPlan } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
import { captureReviewGroupHistory, reviewGroupHistoryRefusal } from '../review/group-history';
import { groupRemovalBudget } from '../review/group-removal-budget';
import { groupReviewFieldBudget } from '../review/group-field-budget';

/** Image fields and their native resource retirement must share the existing owned history scope. */
export function directImageHistory(
  editor: DocxEditorInstance,
  plan: CanonicalEditPlan,
  before: OoxmlPackage,
  after: OoxmlPackage
): <T>(run: () => T) => T {
  const session = editor.surface?.collaborationSession() ?? null;
  if (
    !session ||
    !plan.operations.some((edit) => edit.kind === 'inline-image' || edit.kind === 'inline-remove')
  )
    return (run) => run();
  const refusal = reviewGroupHistoryRefusal(session);
  if (refusal) refuseWrite(refusal, 'unsupported-operation');
  const flags = groupRemovalBudget(before, after);
  const fields = groupReviewFieldBudget(before, after);
  return (run) => captureReviewGroupHistory(session, run, flags, fields);
}
