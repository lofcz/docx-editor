import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { sameRevision } from '../../src/contracts/document';
import { directParagraphTextOps, planParagraphText } from '../../src/diff/paragraph-text';
import type { CanonicalEditPlan } from '../../src/diff/plan';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { docxFixture } from './document';

export const TEXT_DOCUMENT_ID = 'text-plan-test';
export const TEXT_EPOCH = 'mount-1';

export function textPlanDocument() {
  return docxFixture(
    '<w:p><w:r><w:rPr><w:b/><w:customKeep w:val="keep"/></w:rPr><w:t xml:space="preserve">Before </w:t></w:r><w:r><w:t>cat</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> after 😀.</w:t></w:r></w:p><w:p><w:r><w:t>Unchanged tail.</w:t></w:r></w:p>'
  );
}

export function paragraphPlan(editor: DocxEditorInstance, paragraphId: string, target: string) {
  const revision = readEditorRevision(editor, TEXT_DOCUMENT_ID, TEXT_EPOCH);
  const part = editor.surface!.session.part();
  const text = paragraphTextOf(part, paragraphId);
  if (!revision || text === null) throw new Error('Missing native paragraph snapshot');
  const result = planParagraphText({ revision, partName: part.name, paragraphId, text }, target);
  if (!result.success) throw new Error(result.error);
  return result.plan;
}

/** Test-only preflight. A01/A02 own the production executor and its wider validation contract. */
export function applyTextPlan(editor: DocxEditorInstance, plan: CanonicalEditPlan): boolean {
  const revision = readEditorRevision(editor, TEXT_DOCUMENT_ID, TEXT_EPOCH);
  if (!revision || !sameRevision(plan.revision, revision)) return false;
  const surface = editor.surface!;
  const part = surface.session.part();
  for (const edit of plan.operations) {
    if (edit.kind !== 'text' || edit.partName !== part.name) return false;
    const text = paragraphTextOf(part, edit.paragraphId);
    if (text === null || text.slice(edit.start, edit.end) !== edit.before) return false;
  }
  if (!plan.operations.length) return true;
  const ops = plan.operations.flatMap((edit) =>
    edit.kind === 'text' ? directParagraphTextOps(edit) : []
  );
  return !surface.applyAutomationOps(() => ops).rejected;
}
