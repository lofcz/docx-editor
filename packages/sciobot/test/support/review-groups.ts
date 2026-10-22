import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { bodyStoryRoot, paragraphTextOf, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import type { BlockPayload } from '../../src/blocks/payload-schema';
import { planParagraphText } from '../../src/diff/paragraph-text';
import { createEditPlan } from '../../src/diff/plan';
import { parseTargetMarkdown } from '../../src/projection/parse-target';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { readReviewGroups } from '../../src/review/groups';
import { trackSiblingPlan } from '../../src/review/track-blocks';
import { trackTextPlan } from '../../src/review/track-text';
import { atomicListening, atomicVisual } from './atomic-execution';
import { propertyContext } from './property-execution';
import { sourceControl, sourceParagraph } from './target-source-fixture';
import { docxFixture } from './document';

export const groupToken = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, propertyContext.documentId, propertyContext.epoch)!;
export function groupFixture() {
  const p = sourceParagraph,
    label =
      '<w:sdt><w:sdtPr><w:tag w:val="scio:label"/></w:sdtPr><w:sdtContent>' +
      p('Old activity — 10 minutes') +
      '</w:sdtContent></w:sdt>',
    body =
      sourceControl(
        'phase',
        p('The cat sat.') +
          sourceControl(
            'activity',
            label +
              '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:sdt><w:sdtPr><w:tag w:val="ordinary-inner"/></w:sdtPr><w:sdtContent>' +
              p('A dog ran.') +
              '</w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>'
          ) +
          sourceControl('visual', p('Old visual fallback')) +
          sourceControl('audio', p('Old audio fallback'))
      ) + p('Keep tail'),
    read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const payload of [
    {
      id: 'phase',
      type: 'lesson_phase',
      version: 1,
      attributes: { phaseType: 'evocation' },
    } as BlockPayload,
    {
      id: 'activity',
      type: 'lesson_activity',
      version: 1,
      attributes: { name: 'Old activity', duration: 10, resource: 'keep' },
    } as BlockPayload,
    atomicVisual,
    atomicListening,
  ]) {
    const stored = withPayloadRecord(pkg, payload, `Stored ${payload.id} label`);
    if (!stored.ok) throw new Error(stored.reason);
    pkg = stored.pkg;
  }
  return pkg;
}
export function groups(editor: DocxEditorInstance) {
  const result = readReviewGroups(editor.surface!.session.currentPackage());
  if (!result.success) throw new Error(result.error);
  return result.groups;
}
export function groupParagraph(editor: DocxEditorInstance, text: string) {
  const session = editor.surface!.session,
    id = session.paragraphIds().find((id) => paragraphTextOf(session.part(), id) === text);
  if (!id) throw new Error(`Missing group paragraph: ${text}`);
  return id;
}
export function proposeGroupText(editor: DocxEditorInstance, before: string, after: string) {
  const session = editor.surface!.session,
    paragraphId = groupParagraph(editor, before),
    plan = planParagraphText(
      { revision: groupToken(editor), partName: session.part().name, paragraphId, text: before },
      after
    );
  if (!plan.success) throw new Error(plan.error);
  const result = trackTextPlan(editor, plan.plan, propertyContext);
  if (!result.success) throw new Error(result.error);
  return result;
}
export function insertLeadingGroup(editor: DocxEditorInstance) {
  const session = editor.surface!.session,
    body = bodyStoryRoot(session.part())!,
    target = parseTargetMarkdown('New first paragraph');
  if (body.kind === 'textValue' || !target.success)
    throw new Error('Missing group insertion source');
  const plan = createEditPlan(groupToken(editor), [
    {
      kind: 'siblings',
      partName: session.part().name,
      parentId: body.id,
      beforeId: null,
      afterId: body.children[0]!.id,
      removedIds: [],
      inserted: target.target.normalizedAst.children,
    },
  ]);
  if (!plan.success) throw new Error(plan.error);
  const result = trackSiblingPlan(editor, plan.plan, propertyContext);
  if (!result.success) throw new Error(result.error);
  return result;
}
