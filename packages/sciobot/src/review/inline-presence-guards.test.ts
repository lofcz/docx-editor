import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  bodyStoryRoot,
  collectRevisionSites,
  contentControlsIn,
  readOoxmlPackage,
  replaceChildren,
  storyParagraphs,
  withPart,
  writeOoxmlPackage,
  withCustomXmlNode,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createMentionPayload } from '../blocks/mention-codec';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { createBlockPresenceRecord, bindBlockPresenceRecord } from './block-presence-plan';
import { withBlockPresenceRecord } from './block-presence-store';
import type { BlockPresenceRecord } from './block-presence-schema';
import { decodeBlockPresenceRecord, encodeBlockPresenceRecord } from './block-presence-schema';
import { decodePropertyReviewRecord } from './property-record-schema';
import { readPropertyReviewStore } from './property-record-store';
import { PROPERTY_REVIEW_NODE_PREFIX } from '../blocks/payload-schema';
import {
  reviewGroupRevisionReservations,
  withReviewGroupRevisionReservations,
} from './group-allocation';
const attribution = { id: 'inline-proof', author: 'AI', date: '2026-10-24T12:00:00.000Z' };
function fixture() {
  const read = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'));
  if (!read.ok) throw Error(read.reason);
  const before = read.package,
    part = before.parts.get(before.mainDocumentPart)!;
  const paragraphId = storyParagraphs(bodyStoryRoot(part)!)[0]!.id;
  const inserted = applyTreeOp(part, {
    op: 'insertInlineContentControl',
    paragraphId,
    offset: 0,
    tag: 'scio:1:guard-chip',
    text: 'Ada',
    revision: attribution,
  });
  if (!inserted.ok) throw Error(inserted.reason);
  const stored = withPayloadRecord(
    withPart(before, inserted.part),
    createMentionPayload('guard-chip', 'user-7', 'Ada'),
    'Mention'
  );
  if (!stored.ok) throw Error(stored.reason);
  const record = createBlockPresenceRecord(
    before,
    stored.pkg,
    [{ blockId: 'guard-chip', before: { present: false }, after: { present: true } }],
    attribution,
    { nativeInline: true }
  );
  if (!record.success) throw Error(record.error);
  return { before, pkg: stored.pkg, record: record.record };
}

test('inline presence rejects missing, conflicting and forged native ownership without document writes', () => {
  const { pkg, record } = fixture(),
    bytes = writeOoxmlPackage(pkg);
  const data = record.extension!.nativeInline as {
    schemaVersion: number;
    owners: Record<string, unknown>[];
  };
  const owner = data.owners[0]!;
  const altered = (nativeInline: unknown) =>
    ({ ...record, extension: { nativeInline } }) as unknown as BlockPresenceRecord;
  const wrongOwner = (patch: object) =>
    altered({ schemaVersion: 1, owners: [{ ...owner, ...patch }] });
  const revision = owner.revision as Record<string, unknown>;
  const { extension: _profile, ...withoutProfile } = record;
  const records: BlockPresenceRecord[] = [
    withoutProfile,
    altered(null),
    altered([]),
    altered({ schemaVersion: 99, owners: data.owners }),
    altered({ schemaVersion: 1, owners: [] }),
    altered({ schemaVersion: 1, owners: [owner, owner] }),
    wrongOwner({ blockId: 'another-chip' }),
    wrongOwner({ partName: '/word/other.xml' }),
    wrongOwner({ kind: 'del' }),
    wrongOwner({ revision: { ...revision, id: 'missing' } }),
    wrongOwner({ revision: { ...revision, author: 'Other' } }),
    wrongOwner({ revision: { ...revision, date: '2025-01-01T00:00:00Z' } }),
    wrongOwner({ unknown: true }),
    {
      ...record,
      blocks: record.blocks.map((block) => ({
        ...block,
        before: { present: true },
        after: { present: false },
      })),
    },
  ];
  for (const changed of records) {
    expect(bindBlockPresenceRecord(pkg, changed).success).toBe(false);
    expect(withBlockPresenceRecord(pkg, changed).success).toBe(false);
    expect(writeOoxmlPackage(pkg)).toEqual(bytes);
  }
});

test('inline ownership never absorbs foreign text from a containing native revision', () => {
  const { pkg, record } = fixture(),
    part = pkg.parts.get(pkg.mainDocumentPart)!;
  const site = collectRevisionSites(part)[0]!.node,
    paragraph = storyParagraphs(bodyStoryRoot(part)!)[0]!;
  if (paragraph.kind !== 'paragraph') throw Error('Missing source paragraph');
  const original = paragraph.children.find((node) => node.kind === 'run')!;
  let ordinal = 0;
  const copy = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'textValue'
      ? { ...node, id: `foreign:${ordinal++}` }
      : ({ ...node, id: `foreign:${ordinal++}`, children: node.children.map(copy) } as OoxmlNode);
  const changed = replaceChildren(part, site.id, [...site.children, copy(original)]);
  if (!changed.ok) throw Error(JSON.stringify(changed.issues));
  const candidate = withPart(pkg, changed.part);
  expect(contentControlsIn(candidate.parts.get(candidate.mainDocumentPart)!.root)).toHaveLength(1);
  expect(bindBlockPresenceRecord(candidate, record).success).toBe(false);
  expect(withBlockPresenceRecord(candidate, record).success).toBe(false);
});

test('untracked inline roots and malformed creation options cannot acquire native ownership', () => {
  const { before, pkg, record } = fixture();
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  const accepted = applyTreeOp(part, {
    op: 'acceptAllRevisions',
    siteNodeIds: collectRevisionSites(part).map((site) => site.node.id),
  });
  if (!accepted.ok) throw Error(accepted.reason);
  const intents = record.blocks.map(({ blockId, before, after }) => ({ blockId, before, after }));
  expect(
    createBlockPresenceRecord(before, withPart(pkg, accepted.part), intents, attribution, {
      nativeInline: true,
    }).success
  ).toBe(false);
  for (const options of [null, [], 1, '', { nativeInline: false }, { unknown: true }])
    expect(
      createBlockPresenceRecord(
        before,
        pkg,
        record.blocks.map(({ blockId, before, after }) => ({ blockId, before, after })),
        attribution,
        options as unknown as { nativeInline?: true }
      ).success
    ).toBe(false);
});

test('schema 3 reserves inline native IDs after decisions and opaque record tombstones', () => {
  const { pkg, record } = fixture();
  expect(record.schemaVersion).toBe(3);
  const raw = encodeBlockPresenceRecord(record);
  expect(decodePropertyReviewRecord(raw)).toMatchObject({
    status: 'unsupported-version',
    version: 3,
  });
  expect(decodeBlockPresenceRecord(raw).status).toBe('supported');
  expect(decodeBlockPresenceRecord(JSON.stringify({ ...record, schemaVersion: 2 })).status).toBe(
    'invalid'
  );
  const stored = withBlockPresenceRecord(pkg, record);
  if (!stored.success) throw Error(stored.error);
  const part = stored.pkg.parts.get(stored.pkg.mainDocumentPart)!;
  const site = collectRevisionSites(part)[0]!;
  const id = site.node.attributes.find((attribute) => attribute.localName === 'id')!.value;
  const expected = [{ partName: part.name, ids: [id] }];
  expect(reviewGroupRevisionReservations(stored.pkg)).toEqual(expected);
  const accepted = applyTreeOp(part, { op: 'acceptAllRevisions', siteNodeIds: [site.node.id] });
  if (!accepted.ok) throw Error(accepted.reason);
  const decided = withPart(stored.pkg, accepted.part);
  expect(collectRevisionSites(accepted.part)).toEqual([]);
  expect(reviewGroupRevisionReservations(decided)).toEqual(expected);
  const metadata = readPropertyReviewStore(decided);
  if (metadata.status !== 'ready') throw Error('Missing retained proposal store');
  const retired = withCustomXmlNode(decided, metadata.part.partName, {
    id: PROPERTY_REVIEW_NODE_PREFIX + record.id,
    label: 'removed',
    data: raw,
  });
  const read = readPropertyReviewStore(retired);
  // The generic v1 property reader deliberately keeps opaque tombstones visible.
  // An unknown record's label must never hide unsupported ownership.
  expect(read.status === 'ready' && read.records.has(record.id)).toBe(true);
  expect(read.status === 'ready' && read.retainedRecords.has(record.id)).toBe(true);
  expect(reviewGroupRevisionReservations(retired)).toEqual(expected);
  const reopened = readOoxmlPackage(writeOoxmlPackage(retired));
  if (!reopened.ok) throw Error(reopened.reason);
  expect(reviewGroupRevisionReservations(reopened.package)).toEqual(expected);
  const current = reopened.package.parts.get(reopened.package.mainDocumentPart)!;
  const paragraph = storyParagraphs(bodyStoryRoot(current)!)[0]!;
  const next = withReviewGroupRevisionReservations(reopened.package, () =>
    applyTreeOp(current, {
      op: 'insertInlineContentControl',
      paragraphId: paragraph.id,
      offset: 0,
      tag: 'test:next',
      text: 'Next',
      revision: attribution,
    })
  );
  if ('success' in next || !next.ok) throw Error('Reserved native insertion refused');
  const fresh = collectRevisionSites(next.part);
  expect(fresh).toHaveLength(1);
  expect(
    fresh[0]!.node.attributes.find((attribute) => attribute.localName === 'id')!.value
  ).not.toBe(id);
});
