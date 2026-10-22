import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { planParagraphText } from '../diff/paragraph-text';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { applyDirectPlan } from '../ai/apply-direct';
import type { NativeEditFailure } from '../contracts/document';
import { planListMembership } from '../diff/lists';
import { planParagraphFormat } from '../diff/paragraph-format';
import { createEditPlan, type CanonicalEdit } from '../diff/plan';

export type TextBlockType = 'p' | 'h1' | 'h2' | 'h3' | 'blockquote';
export const BLOCK_PARAGRAPH_STYLES: Readonly<Record<TextBlockType, string>> = Object.freeze({
  p: 'Normal',
  h1: 'Heading1',
  h2: 'Heading2',
  h3: 'Heading3',
  blockquote: 'Quote',
});
export interface BlockCommandContext {
  readonly documentId: string;
  readonly epoch: string;
}
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

/** One native undo unit changes selected paragraphs and removes their previous list membership. */
export function setBlockType(
  editor: DocxEditorInstance,
  type: TextBlockType,
  context: BlockCommandContext,
  consumePrefix?: string
) {
  const surface = editor.surface;
  if (!surface) return fail('The native editor is not attached.');
  surface.flushPendingInput();
  const style = BLOCK_PARAGRAPH_STYLES[type];
  if (
    !style ||
    !surface.session
      .documentStyles()
      .some((entry) => entry.type === 'paragraph' && entry.styleId === style)
  )
    return fail('The document does not define this paragraph style.');
  const session = surface.session;
  const pkg = session.currentPackage();
  const part = session.part();
  if (part.name !== pkg.mainDocumentPart) return fail('Block controls require the body story.');
  const selection = surface.state().selection;
  const all = session.paragraphIds();
  const a = all.indexOf(selection.anchor.paragraphId),
    h = all.indexOf(selection.head.paragraphId);
  if (a < 0 || h < 0) return fail('The native selection has no paragraph.');
  const end = a <= h ? selection.head : selection.anchor;
  const selected = all.slice(
    Math.min(a, h),
    Math.max(a, h) + (a !== h && end.offset === 0 ? 0 : 1)
  );
  const revision = { ...context, revision: session.packageRevision() };
  const lists = planListMembership(
    { pkg, partName: part.name, revision },
    selected.map((paragraphId) => ({ paragraphId, target: { kind: 'none' } }))
  );
  if (!lists.success) return lists;
  const operations: CanonicalEdit[] = [...lists.plan.operations];
  if (consumePrefix !== undefined) {
    if (
      selected.length !== 1 ||
      selection.anchor.offset !== selection.head.offset ||
      selection.head.offset !== consumePrefix.length ||
      paragraphTextOf(part, selected[0]!) !== consumePrefix
    )
      return fail('The input prefix no longer matches the native selection.');
    const text = planParagraphText(
      { revision, partName: part.name, paragraphId: selected[0]!, text: consumePrefix },
      ''
    );
    if (!text.success) return text;
    operations.push(...text.plan.operations);
  }
  for (const paragraphId of selected) {
    const format = planParagraphFormat(
      { part, paragraphId, revision },
      { properties: [{ localName: 'pStyle', attributes: { val: style } }] }
    );
    if (!format.success) return format;
    operations.push(...format.plan.operations);
  }
  const plan = createEditPlan(revision, operations);
  if (!plan.success) return plan;
  return applyDirectPlan(editor, plan.plan, context);
}
