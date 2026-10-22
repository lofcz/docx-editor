import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, writeOoxmlPackage, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { planDomainProperties, type DomainPropertyTarget } from '../../src/diff/domain-properties';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../../src/review/native-module';
import { propertyFallbackParagraph } from '../../src/review/property-fallback';
import { trackPropertyPlan } from '../../src/review/track-properties';
import { propertyRecordFixture, propertyRecordRevision } from './property-record';

export const propertyContext = {
  documentId: propertyRecordRevision.documentId,
  epoch: propertyRecordRevision.epoch,
  author: 'ScioBot',
  date: '2026-10-19T09:00:00Z',
  proposalId: 'property-proposal',
};
export function propertyEditor(pkg: OoxmlPackage = propertyRecordFixture(), review = true) {
  return createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });
}
export const propertyToken = (editor: DocxEditorInstance, epoch = propertyContext.epoch) =>
  readEditorRevision(editor, propertyContext.documentId, epoch)!;
export function nativePropertyPlan(
  editor: DocxEditorInstance,
  target: DomainPropertyTarget,
  blockId = 'activity'
) {
  const result = planDomainProperties(
    {
      pkg: editor.surface!.session.currentPackage(),
      revision: propertyToken(editor),
      blockId,
    },
    target
  );
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
export function propertyLabel(editor: DocxEditorInstance, blockId = 'activity') {
  const session = editor.surface!.session,
    owner = buildBlockIndex(session.currentPackage()).byId.get(blockId)!,
    paragraph = propertyFallbackParagraph(session.part(), owner.controlId);
  if (!paragraph) throw new Error('Missing native generated label');
  return paragraph;
}
export function executeProperties(
  editor: DocxEditorInstance,
  target: DomainPropertyTarget = { set: { name: 'New activity', duration: 15 } },
  after = 'New activity — 15 minutes',
  proposalId = propertyContext.proposalId,
  blockId = 'activity'
) {
  const session = editor.surface!.session,
    paragraph = propertyLabel(editor, blockId);
  return trackPropertyPlan(
    editor,
    nativePropertyPlan(editor, target, blockId),
    [{ blockId, before: paragraphTextOf(session.part(), paragraph.id)!, after }],
    { ...propertyContext, proposalId }
  );
}
