import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { bodyStoryRoot, escapeXml, type TreeDocOp } from '@docx-editor.dev/core/store';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { parseTargetMarkdown } from '../../src/projection/parse-target';
import { planTargetDocument } from '../../src/diff/plan-target';
import { createEditPlan, type CanonicalEdit } from '../../src/diff/plan';
import { readCommittedView } from '../../src/review/committed-view';
import { readWorkingView } from '../../src/review/working-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../../src/review/native-module';
import type { NativeSiblingProposal } from '../../src/review/track-blocks';
import { docxFixture } from './document';

export const blockContext = {
  documentId: 'sibling-test',
  epoch: 'mount-1',
  author: 'ScioBot',
  date: '2026-10-19T09:00:00Z',
};
export const blockP = (text: string) =>
  `<w:p keep:flag="paragraph" xmlns:keep="urn:keep"><w:pPr><keep:metadata/></w:pPr><w:r keep:flag="run"><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
export function blockEditor(body: string, review = true) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docxFixture(body),
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });
  return {
    editor,
    destroy() {
      editor.destroy();
      container.remove();
    },
  };
}
export const blockToken = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, blockContext.documentId, blockContext.epoch)!;
export function blockAst(markdown: string) {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return parsed.target.normalizedAst;
}
export async function blockViews(editor: DocxEditorInstance) {
  const token = blockToken(editor),
    [committed, working] = await Promise.all([
      readCommittedView(editor.surface!.session, token),
      readWorkingView(editor.surface!.session, token),
    ]);
  if (!committed.success) throw new Error(committed.error);
  if (!working.success) throw new Error(working.error);
  return [committed.view.projection.text, working.view.projection.text];
}
export async function targetBlockPlan(editor: DocxEditorInstance, target: string) {
  const token = blockToken(editor),
    working = await readWorkingView(editor.surface!.session, token);
  if (!working.success) throw new Error(working.error);
  const planned = planTargetDocument(
    {
      revision: token,
      pkg: editor.surface!.session.currentPackage(),
      projection: working.view.projection,
    },
    target
  );
  if (!planned.success) throw new Error(planned.error);
  return planned.plan;
}
export function manualBlockPlan(
  editor: DocxEditorInstance,
  start: number,
  count: number,
  target = ''
) {
  const part = editor.surface!.session.part(),
    body = bodyStoryRoot(part)!;
  if (body.kind === 'textValue') throw new Error('Missing native body');
  const nodes = body.children.filter((node) => node.kind === 'paragraph');
  const edit: CanonicalEdit = {
    kind: 'siblings',
    partName: part.name,
    parentId: body.id,
    beforeId: start === 0 ? null : nodes[start - 1]!.id,
    afterId: nodes[start + count]?.id ?? null,
    removedIds: nodes.slice(start, start + count).map((node) => node.id),
    inserted: blockAst(target)
      .children as readonly import('../../src/projection/parse-target').TargetAstNode[],
  };
  const made = createEditPlan(blockToken(editor), [edit]);
  if (!made.success) throw new Error(made.error);
  return made.plan;
}
export function decideBlocks(
  editor: DocxEditorInstance,
  proposals: readonly NativeSiblingProposal[],
  decision: 'acceptRevision' | 'rejectRevision'
) {
  const ops: TreeDocOp[] = proposals.flatMap((proposal) =>
    proposal.addresses.map((revision) => ({ op: decision, revision }))
  );
  const result = editor.surface!.applyAutomationOps(
    () => ops,
    { kind: 'body' },
    undefined,
    true,
    true
  );
  if (result.rejected || !result.committed) throw new Error(JSON.stringify(result));
}
