import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  withInitializedReviewGroupStore,
  withReviewGroupRecord,
} from '../../src/review/group-record-store';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { planDomainProperties } from '../../src/diff/domain-properties';
import { createEditPlan } from '../../src/diff/plan';
import { propertyFallbackParagraph } from '../../src/review/property-fallback';
import { trackPropertyPlan } from '../../src/review/track-properties';
import { acceptReviewGroup } from '../../src/review/accept-group';
import { groups, groupToken } from './review-groups';
import { propertyContext } from './property-execution';
import { sourceControl, sourceParagraph } from './target-source-fixture';
import { docxFixture, W } from './document';

export const decisionBlocks = ['first', 'middle', 'last'] as const;
export function groupDecisionFixture() {
  const read = readOoxmlPackage(
    docxFixture(
      sourceControl(
        'phase',
        decisionBlocks
          .map((id) =>
            sourceControl(
              id,
              '<w:sdt><w:sdtPr><w:tag w:val="scio:label"/></w:sdtPr><w:sdtContent>' +
                sourceParagraph(`Old ${id}`) +
                '</w:sdtContent></w:sdt>' +
                sourceParagraph(`Body ${id}`)
            )
          )
          .join('')
      ) + sourceParagraph('Keep tail')
    )
  );
  if (!read.ok) throw new Error(read.reason);
  let pkg = read.package;
  for (const id of ['phase', ...decisionBlocks]) {
    const written = withPayloadRecord(
      pkg,
      {
        version: 1,
        id,
        type: id === 'phase' ? 'lesson_phase' : 'lesson_activity',
        attributes:
          id === 'phase'
            ? { phaseType: 'evocation' }
            : { name: `Old ${id}`, duration: 10, resource: 'keep' },
        extension: { opaque: [1, null, 'keep'] },
      },
      `Stored ${id}`
    );
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  return pkg;
}
export function proposeThreeProperties(editor: DocxEditorInstance, metadataMiddle = false) {
  const pkg = editor.surface!.session.currentPackage(),
    revision = groupToken(editor),
    operations = decisionBlocks.flatMap((blockId) => {
      const result = planDomainProperties(
        { pkg, revision, blockId },
        { set: { name: `New ${blockId}`, duration: 15 } }
      );
      if (!result.success) throw new Error(result.error);
      return result.plan.operations;
    }),
    plan = createEditPlan(revision, operations);
  if (!plan.success) throw new Error(plan.error);
  const result = trackPropertyPlan(
    editor,
    plan.plan,
    decisionBlocks.map((blockId) => ({
      blockId,
      before: `Old ${blockId}`,
      after: metadataMiddle && blockId === 'middle' ? `Old ${blockId}` : `New ${blockId}`,
    })),
    { ...propertyContext, proposalId: 'three-blocks' }
  );
  if (!result.success) throw new Error(result.error);
  return result;
}
export function acceptBlockGroup(editor: DocxEditorInstance, blockId = 'middle') {
  const group = groups(editor).find((group) =>
    group.properties.some((member) => member.blockId === blockId)
  );
  if (!group) throw new Error('Missing block group ' + blockId);
  return acceptReviewGroup(editor, group.id, groupToken(editor), propertyContext);
}
export function blockLabel(editor: DocxEditorInstance, blockId: string) {
  const session = editor.surface!.session,
    block = buildBlockIndex(session.currentPackage()).byId.get(blockId)!,
    paragraph = propertyFallbackParagraph(session.part(), block.controlId)!;
  return paragraphTextOf(session.part(), paragraph.id);
}

export function crossStoryGroupFixture() {
  const revision = '<w:ins w:id="7" w:author="Teacher"><w:r><w:t>Pending</w:t></w:r></w:ins>',
    entries = unzipSync(
      docxFixture(
        '<w:p><w:r><w:t xml:space="preserve">Body </w:t></w:r>' +
          revision +
          '</w:p><w:p><w:ins w:id="8" w:author="Peer"><w:r><w:t>Other</w:t></w:r></w:ins></w:p><w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr>',
        {
          relationships:
            '<Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
        }
      )
    );
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
    )
  );
  entries['word/header1.xml'] = strToU8(`<w:hdr xmlns:w="${W}"><w:p>${revision}</w:p></w:hdr>`);
  const read = readOoxmlPackage(zipSync(entries));
  if (!read.ok) throw new Error(read.reason);
  const initialized = withInitializedReviewGroupStore(read.package);
  if (!initialized.ok) throw new Error(initialized.reason);
  const stored = withReviewGroupRecord(initialized.pkg, {
    schemaVersion: 1,
    id: 'rg_cross_story',
    members: ['/word/document.xml', '/word/header1.xml'].map((partName) => ({
      partName,
      revision: { id: '7', author: 'Teacher' },
    })),
  });
  if (!stored.ok) throw new Error(stored.reason);
  return stored.pkg;
}
