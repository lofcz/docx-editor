import { expect, test } from 'bun:test';
import { insertChildren, readOoxmlPart, withPart } from '@docx-editor.dev/core/store';
import { presenceAttribution, presencePackage } from '../../test/support/block-presence';
import { sourceControl, sourceParagraph } from '../../test/support/target-source-fixture';
import {
  propertyRecord,
  propertyRecordFixture,
  propertyRecordRevision,
  packageWithPropertyRecord,
} from '../../test/support/property-record';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { createBlockPresenceRecord } from './block-presence-plan';
import { withBlockPresenceRecord } from './block-presence-store';
import { readApplicationProposals } from './application-proposals';
import { readPropertyReviewStore } from './property-record-store';
import { projectPendingProperties } from './property-projection';
import {
  withInitializedReviewGroupStore,
  withPayloadGroupDecisionRecord,
} from './group-record-store';

function mixed() {
  const base = propertyRecordFixture();
  const part = base.parts.get(base.mainDocumentPart)!;
  const seed = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:opaque="urn:opaque">${'<w:body>' + sourceControl('outer', sourceControl('inner', sourceParagraph('Visual'))) + '</w:body>'}</w:document>`,
    {
      name: '/presence-causal-seed.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!seed.ok) throw Error('Missing seed');
  const body = part.root.children.find((node) => node.kind === 'body')!;
  const seedBody = seed.part.root.children.find((node) => node.kind === 'body')!;
  if (seedBody.kind === 'textValue') throw Error('Invalid body');
  const inserted = insertChildren(part, body.id, 0, seedBody.children);
  if (!inserted.ok) throw Error('Insertion failed');
  let pkg = withPart(base, inserted.part);
  for (const block of buildBlockIndex(presencePackage()).byId.values()) {
    const stored = withPayloadRecord(pkg, block.payload, block.id + ' label');
    if (!stored.ok) throw Error(stored.reason);
    pkg = stored.pkg;
  }
  const first = { ...propertyRecord('first'), nativeRevisions: [] };
  pkg = packageWithPropertyRecord(pkg, first);
  const record = createBlockPresenceRecord(
    base,
    pkg,
    [{ blockId: 'outer', before: { present: false }, after: { present: true } }],
    presenceAttribution
  );
  if (!record.success) throw Error(record.error);
  const stored = withBlockPresenceRecord(pkg, { ...record.record, predecessors: ['first'] });
  if (!stored.success) throw Error(stored.error);
  pkg = packageWithPropertyRecord(stored.pkg, {
    ...propertyRecord('last'),
    nativeRevisions: [],
    predecessors: [presenceAttribution.id],
    blocks: [
      {
        blockId: 'activity',
        type: 'lesson_activity',
        changes: [
          {
            name: 'duration',
            before: { present: true, value: 15 },
            after: { present: true, value: 20 },
          },
        ],
      },
    ],
  });
  return pkg;
}
test('mixed formats retain transitive property causality through a presence record without rewriting stored bytes', () => {
  const pkg = mixed();
  const before = readPropertyReviewStore(pkg);
  const read = readApplicationProposals(pkg, propertyRecordRevision);
  if (!read.success) throw Error(read.error);
  expect(read.properties.map((proposal) => proposal.id)).toEqual(['first', 'last']);
  expect(read.presence.map((record) => record.id)).toEqual([presenceAttribution.id]);
  for (const kind of ['working', 'committed'] as const) {
    const projected = projectPendingProperties(pkg, propertyRecordRevision, read.properties, kind);
    if (!projected.success) throw Error(projected.error);
    expect(buildBlockIndex(projected.pkg).byId.get('activity')!.payload.attributes.duration).toBe(
      kind === 'working' ? 20 : 10
    );
  }
  expect(readPropertyReviewStore(pkg)).toEqual(before);
});
test('a resolved presence record still connects property causality', () => {
  let pkg = mixed();
  const initialized = withInitializedReviewGroupStore(pkg);
  if (!initialized.ok) throw Error(initialized.reason);
  const decided = withPayloadGroupDecisionRecord(initialized.pkg, {
    schemaVersion: 2,
    id: 'rg_bridge',
    proposalId: presenceAttribution.id,
    blockId: 'outer',
    decision: 'accept',
  });
  if (!decided.ok) throw Error(decided.reason);
  pkg = decided.pkg;
  const read = readApplicationProposals(pkg, propertyRecordRevision);
  if (!read.success) throw Error(read.error);
  expect(read.properties.map((proposal) => proposal.id)).toEqual(['first', 'last']);
  expect(read.presence).toHaveLength(0);
});
