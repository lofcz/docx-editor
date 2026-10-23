import { readOoxmlPart, insertChildren, withPart } from '@docx-editor.dev/core/store';
import { propertyRecordFixture } from './property-record';
import { sourceControl, sourceParagraph } from './target-source-fixture';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import { createBlockPresenceRecord } from '../../src/review/block-presence-plan';
import { withBlockPresenceRecord } from '../../src/review/block-presence-store';
import { captureOwnedProposalContentProofs } from '../../src/review/owned-proposal-content';
import { readReviewGroups } from '../../src/review/groups';
import { presenceAttribution, presencePackage } from './block-presence';
import { propertyToken } from './property-execution';

export function presenceReviewFixture() {
  const pkg = presencePackage();
  const inner = buildBlockIndex(pkg).byId.get('inner')!;
  const changed = withPayloadRecord(pkg, { ...inner.payload, type: 'lesson_visual' }, 'inner label');
  if (!changed.ok) throw Error(changed.reason);
  return changed.pkg;
}
export function installPresenceOwnership(editor: DocxEditorInstance, remove: boolean) {
  const baseline = editor.surface!.session.currentPackage();
  const prepared = createBlockPresenceRecord(remove ? baseline : presencePackage(false), baseline,
    [{ blockId: 'outer', before: { present: remove }, after: { present: !remove } }], presenceAttribution);
  if (!prepared.success) throw Error(prepared.error);
  const stored = editor.surface!.applyAutomationOps(() => [], { kind: 'body' }, [working => {
    const result = withBlockPresenceRecord(working, prepared.record);
    if (!result.success) throw Error(result.error);
    return result.pkg;
  }], true, true, 'after');
  if (stored.rejected) throw Error(String(stored.reason));
  const read = readReviewGroups(editor.surface!.session.currentPackage());
  if (!read.success) throw Error(read.error);
  const ownership = read.groups.filter(group => group.properties.some(member => member.proposalId === presenceAttribution.id)).map(group => ({
    groupId: group.id,
    native: group.native.map(member => ({ partName: member.partName, addresses: [member.address] })),
    properties: group.properties,
  }));
  const proofs = captureOwnedProposalContentProofs(editor.surface!.session.currentPackage(), propertyToken(editor), ownership);
  if (!proofs.success) throw Error(proofs.error);
  return { baseline, ownership, contentProofs: proofs.contentProofs };
}

/** Independent application control beside the existing editable lesson activity. */
export function presencePropertyFixture() {
  const base = propertyRecordFixture(), part = base.parts.get(base.mainDocumentPart)!;
  const seed = readOoxmlPart(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${sourceControl('outer', sourceControl('inner', sourceParagraph('Opaque child body')))}</w:body></w:document>`,
    { name: '/presence-property-seed.xml', contentType: part.contentType });
  if (!seed.ok) throw Error('Missing presence seed');
  const source = seed.part.root.children.find(node => node.kind === 'body')!;
  const body = part.root.children.find(node => node.kind === 'body')!;
  if (source.kind === 'textValue') throw Error('Invalid presence seed');
  const inserted = insertChildren(part, body.id, 0, source.children);
  if (!inserted.ok) throw Error('Presence insertion failed');
  let pkg = withPart(base, inserted.part);
  for (const block of buildBlockIndex(presenceReviewFixture()).byId.values()) {
    const stored = withPayloadRecord(pkg, block.payload, block.id + ' label');
    if (!stored.ok) throw Error(stored.reason);
    pkg = stored.pkg;
  }
  return pkg;
}
