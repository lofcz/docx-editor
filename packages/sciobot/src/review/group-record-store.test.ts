import { expect, test } from 'bun:test';
import {
  customXmlNodes,
  escapeXml,
  readOoxmlPackage,
  readOoxmlPart,
  withCustomXmlDataPart,
  withCustomXmlNode,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { prepareReviewGroupDocument } from '../native/prepare-review-groups';
import {
  decodeReviewGroupRecord,
  encodeReviewGroupRecord,
  REVIEW_GROUP_RECORD_LIMITS,
  type ReviewGroupRecord,
} from './group-record-schema';
import {
  readReviewGroupStore,
  REVIEW_GROUP_STORE_LIMITS,
  REVIEW_GROUP_STORE_NAMESPACE,
  REVIEW_GROUP_STORE_ROOT,
  withInitializedReviewGroupStore,
  withReviewGroupRecord,
  withoutReviewGroupRecord,
} from './group-record-store';

const record = (id = 'rg_one'): ReviewGroupRecord => ({
  schemaVersion: 1,
  id,
  members: [
    {
      partName: '/word/document.xml',
      revision: { id: '17', author: 'ScioBot', date: '2026-10-20T10:00:00.000Z' },
    },
    { partName: '/word/document.xml', revision: { id: '18', author: 'ScioBot' } },
  ],
  extension: { opaque: [null, false, { keep: true }] },
});
function fixture() {
  const read = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Original</w:t></w:r></w:p>'));
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}
function initialized(pkg = fixture()) {
  const result = withInitializedReviewGroupStore(pkg);
  if (!result.ok) throw new Error(result.reason);
  return result.pkg;
}
function store(pkg: OoxmlPackage) {
  const read = readReviewGroupStore(pkg);
  if (read.status !== 'ready') throw new Error('Missing group store');
  return read;
}
function saved(pkg: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}
const node = (id: string, raw: string, extra = '') =>
  `<node id="${id}"><label/><data>${escapeXml(raw)}</data>${extra}</node>`;
function replaceStore(pkg: OoxmlPackage, children: string, root = REVIEW_GROUP_STORE_ROOT) {
  const part = pkg.parts.get(store(pkg).part.partName)!,
    read = readOoxmlPart(`<${root} xmlns="${REVIEW_GROUP_STORE_NAMESPACE}">${children}</${root}>`, {
      name: part.name,
      contentType: part.contentType,
    });
  if (!read.ok) throw new Error(read.reason);
  return withPart(pkg, read.part);
}

test('the independent group store initializes once and keeps native parts and unrelated stores intact', () => {
  const original = fixture(),
    foreign = withCustomXmlDataPart(original, original.mainDocumentPart, 'urn:foreign', 'foreign');
  if (!foreign.part) throw new Error('Missing foreign store');
  const pkg = withCustomXmlNode(foreign.pkg, foreign.part.partName, {
      id: 'keep',
      label: 'Original',
      data: ' exact opaque data ',
    }),
    before = writeOoxmlPackage(pkg),
    candidate = initialized(pkg);
  expect(readReviewGroupStore(pkg)).toEqual({ status: 'absent' });
  expect(candidate.parts.get(pkg.mainDocumentPart)).toBe(pkg.parts.get(pkg.mainDocumentPart));
  expect(candidate.parts.get(foreign.part.partName)).toBe(pkg.parts.get(foreign.part.partName));
  expect(store(candidate).part.partName).not.toBe(foreign.part.partName);
  expect(store(candidate).records.size).toBe(0);
  expect(initialized(candidate)).toBe(candidate);
  expect(store(saved(candidate)).records.size).toBe(0);
  expect(writeOoxmlPackage(pkg)).toEqual(before);
});

test('group records preserve exact Word membership, independent identities, removal receipts, and save/reopen', () => {
  const pkg = initialized(),
    first = withReviewGroupRecord(pkg, record());
  if (!first.ok) throw new Error(first.reason);
  const second = withReviewGroupRecord(first.pkg, record('rg_peer'));
  if (!second.ok) throw new Error(second.reason);
  expect([...store(second.pkg).records.keys()].sort()).toEqual(['rg_one', 'rg_peer']);
  expect(store(saved(second.pkg)).records.get('rg_one')!.decoded).toEqual({
    status: 'supported',
    value: record(),
  });
  expect(withReviewGroupRecord(second.pkg, record())).toEqual({ ok: true, pkg: second.pkg });
  expect(
    withReviewGroupRecord(second.pkg, { ...record(), extension: { changed: true } })
  ).toMatchObject({ ok: false, pkg: second.pkg, reason: 'review-group-record-is-immutable' });
  const removed = withoutReviewGroupRecord(second.pkg, 'rg_one');
  if (!removed.ok) throw new Error(removed.reason);
  expect([...store(removed.pkg).records.keys()]).toEqual(['rg_peer']);
  expect([...store(removed.pkg).removedIds]).toEqual(['rg_one']);
  expect(withReviewGroupRecord(removed.pkg, record())).toMatchObject({
    ok: false,
    reason: 'review-group-id-reserved',
  });
  expect(
    customXmlNodes(removed.pkg, store(removed.pkg).part.partName).find(
      (node) => node.id === 'rg_one'
    )!.data
  ).toBe(encodeReviewGroupRecord(record()));
  expect(withoutReviewGroupRecord(removed.pkg, 'rg_missing')).toEqual({
    ok: true,
    pkg: removed.pkg,
  });
  expect(second.pkg.parts.get(pkg.mainDocumentPart)).toBe(pkg.parts.get(pkg.mainDocumentPart));
});

test('future records retain their raw bytes and cannot be changed or removed by a known writer', () => {
  const raw = '{ "schemaVersion": 3, "id": "rg_future", "unknown": [1, " keep "] }',
    pkg = replaceStore(initialized(), node('rg_future', raw));
  expect(store(pkg).records.get('rg_future')!.decoded).toEqual({
    status: 'unsupported-version',
    id: 'rg_future',
    version: 3,
    raw,
  });
  expect(withReviewGroupRecord(pkg, record('rg_future'))).toMatchObject({
    ok: false,
    pkg,
    reason: 'unsupported-review-group-version',
  });
  expect(withoutReviewGroupRecord(pkg, 'rg_future')).toMatchObject({
    ok: false,
    pkg,
    reason: 'unsupported-review-group-version',
  });
  expect(initialized(pkg)).toBe(pkg);
  const peer = withReviewGroupRecord(pkg, record('rg_known'));
  if (!peer.ok) throw new Error(peer.reason);
  expect(store(saved(peer.pkg)).records.get('rg_future')!.raw).toBe(raw);
});

test('membership decoding is detached, bounded, and explicit about empty grid authors and optional dates', () => {
  const source = {
      ...record(),
      members: [{ partName: '/word/document.xml', revision: { id: '3', author: '' } }],
    },
    raw = encodeReviewGroupRecord(source),
    decoded = decodeReviewGroupRecord(raw);
  if (decoded.status !== 'supported') throw new Error('Missing decoded group');
  source.members[0]!.revision.id = 'changed';
  expect(decoded.value.members[0]!.revision).toEqual({ id: '3', author: '' });
  expect(Object.isFrozen(decoded.value.members[0]!.revision)).toBe(true);
  expect(Object.hasOwn(decoded.value.members[0]!.revision, 'date')).toBe(false);
  const invalid: unknown[] = [
    { ...record(), id: 'bad/id' },
    { ...record(), id: 'rg_' + 'x'.repeat(97) },
    { ...record(), nodeId: 'load-specific' },
    { ...record(), members: [] },
    { ...record(), members: [record().members[0], record().members[0]] },
    {
      ...record(),
      members: [{ partName: '/word/../document.xml', revision: { id: '1', author: 'A' } }],
    },
    { ...record(), members: [{ partName: '/word/document.xml', revision: { id: '1' } }] },
    {
      ...record(),
      members: [
        {
          partName: '/word/document.xml',
          nodeId: 'load-specific',
          revision: { id: '1', author: 'A' },
        },
      ],
    },
    {
      ...record(),
      members: [{ partName: '/word/document.xml', revision: { id: '1', author: 'A', date: null } }],
    },
    { ...record(), extension: [] },
  ];
  for (const value of invalid)
    expect(() => encodeReviewGroupRecord(value as ReviewGroupRecord)).toThrow();
  expect(decodeReviewGroupRecord('invalid')).toEqual({ status: 'invalid', reason: 'invalid-json' });
  expect(decodeReviewGroupRecord('x'.repeat(REVIEW_GROUP_RECORD_LIMITS.maxBytes + 1))).toEqual({
    status: 'invalid',
    reason: 'byte-limit',
  });
  expect(() =>
    encodeReviewGroupRecord({
      ...record(),
      members: Array.from({ length: REVIEW_GROUP_RECORD_LIMITS.maxMembers + 1 }, (_, id) => ({
        partName: '/word/document.xml',
        revision: { id: String(id), author: 'A' },
      })),
    })
  ).toThrow('invalid-review-group-members');
});

test('record writes refuse live getters, cycles, and uninitialized storage without mutating the source', () => {
  let calls = 0;
  const getter = {
      ...record(),
      get members(): ReviewGroupRecord['members'] {
        calls++;
        throw new Error('executed');
      },
    },
    cycle: Record<string, unknown> = {};
  cycle.self = cycle;
  expect(() => encodeReviewGroupRecord(getter)).toThrow('non-json-property');
  expect(calls).toBe(0);
  expect(() => encodeReviewGroupRecord({ ...record(), extension: cycle as never })).toThrow(
    'cyclic-payload'
  );
  const pkg = fixture(),
    bytes = writeOoxmlPackage(pkg);
  expect(withReviewGroupRecord(pkg, record())).toMatchObject({
    ok: false,
    pkg,
    reason: 'review-group-store-initialization-required',
  });
  expect(writeOoxmlPackage(pkg)).toEqual(bytes);
});

test('malformed store nodes, duplicate ownership, ambiguous fields, and record overflow refuse atomically', () => {
  const pkg = initialized(),
    raw = encodeReviewGroupRecord(record()),
    variants = [
      replaceStore(pkg, node('rg_one', raw), 'wrong'),
      replaceStore(pkg, node('rg_one', raw) + node('rg_one', raw)),
      replaceStore(pkg, '<node><label/><data>unknown</data></node>'),
      replaceStore(pkg, node('rg_one', raw, '<data>duplicate</data>')),
      replaceStore(
        pkg,
        `<node id="rg_one"><label/><data><value>${escapeXml(raw)}</value></data></node>`
      ),
      replaceStore(pkg, node('rg_other', raw)),
      replaceStore(
        pkg,
        Array.from({ length: REVIEW_GROUP_STORE_LIMITS.maxRecords + 1 }, (_, index) =>
          node('rg_' + index, '{}')
        ).join('')
      ),
    ];
  for (const candidate of variants) {
    const bytes = writeOoxmlPackage(candidate);
    expect(readReviewGroupStore(candidate).status).toBe('invalid');
    expect(withInitializedReviewGroupStore(candidate)).toMatchObject({ ok: false, pkg: candidate });
    expect(withReviewGroupRecord(candidate, record('rg_new'))).toMatchObject({
      ok: false,
      pkg: candidate,
    });
    expect(writeOoxmlPackage(candidate)).toEqual(bytes);
  }
  const inert = replaceStore(pkg, '<node/>');
  expect(store(inert).records.size).toBe(0);
});

test('group-only byte initialization is idempotent and leaves all original native part bytes intact', () => {
  const input = docxFixture(
      '<w:p><w:ins w:id="5" w:author="Teacher"><w:r><w:t>Pending</w:t></w:r></w:ins></w:p>'
    ),
    prepared = prepareReviewGroupDocument(input);
  if (!prepared.success) throw new Error(prepared.error);
  expect(prepared.changed).toBe(true);
  const before = readOoxmlPackage(input),
    after = readOoxmlPackage(prepared.document);
  if (!before.ok || !after.ok) throw new Error('Missing initialized package');
  expect(after.package.partBytes.get(before.package.mainDocumentPart)).toEqual(
    before.package.partBytes.get(before.package.mainDocumentPart)
  );
  expect(store(after.package).records.size).toBe(0);
  const repeated = prepareReviewGroupDocument(prepared.document);
  expect(repeated).toEqual({ success: true, changed: false, document: prepared.document });
  if (!repeated.success) throw new Error(repeated.error);
  expect(repeated.document).toBe(prepared.document);
});

test('store slot and aggregate data limits reject oversized native metadata before a write', () => {
  const pkg = initialized(),
    slots = replaceStore(pkg, '<node/>'.repeat(REVIEW_GROUP_STORE_LIMITS.maxNodeSlots + 1));
  expect(readReviewGroupStore(slots)).toEqual({
    status: 'invalid',
    reason: 'review-group-store-node-limit',
  });
  const count = replaceStore(
    pkg,
    Array.from({ length: REVIEW_GROUP_STORE_LIMITS.maxRecords + 1 }, (_, id) =>
      node('rg_' + id, '{}')
    ).join('')
  );
  expect(readReviewGroupStore(count)).toEqual({
    status: 'invalid',
    reason: 'review-group-store-count-limit',
  });
  let candidate = pkg;
  const partName = store(pkg).part.partName,
    pad = 'x'.repeat(1_850_000);
  for (let id = 0; id < 5; id++)
    candidate = withCustomXmlNode(candidate, partName, {
      id: 'rg_big_' + id,
      label: '',
      data: JSON.stringify({ schemaVersion: 3, id: 'rg_big_' + id, pad }),
    });
  expect(readReviewGroupStore(candidate)).toEqual({
    status: 'invalid',
    reason: 'review-group-store-byte-limit',
  });
  expect(withReviewGroupRecord(candidate, record())).toMatchObject({
    ok: false,
    pkg: candidate,
    reason: 'review-group-store-byte-limit',
  });
  expect(candidate.parts.get(pkg.mainDocumentPart)).toBe(pkg.parts.get(pkg.mainDocumentPart));
});

test('duplicate store relationships, wrong roots, and invalid byte inputs never initialize over existing data', () => {
  const pkg = initialized(),
    other = withCustomXmlDataPart(pkg, pkg.mainDocumentPart, 'urn:other', 'groups');
  if (!other.part) throw new Error('Missing second store');
  const read = readOoxmlPart(`<groups xmlns="${REVIEW_GROUP_STORE_NAMESPACE}"/>`, {
    name: other.part.partName,
    contentType: other.pkg.parts.get(other.part.partName)!.contentType,
  });
  if (!read.ok) throw new Error(read.reason);
  const duplicate = withPart(other.pkg, read.part);
  expect(readReviewGroupStore(duplicate)).toEqual({
    status: 'invalid',
    reason: 'duplicate-review-group-store',
  });
  expect(withInitializedReviewGroupStore(duplicate)).toMatchObject({ ok: false, pkg: duplicate });
  const wrongRoot = replaceStore(pkg, '', 'wrongRoot'),
    bytes = writeOoxmlPackage(wrongRoot);
  expect(prepareReviewGroupDocument(bytes)).toMatchObject({
    success: false,
    code: 'invalid-document',
    error: 'invalid-review-group-store-root',
  });
  expect(writeOoxmlPackage(wrongRoot)).toEqual(bytes);
  expect(prepareReviewGroupDocument(new Uint8Array([1, 2, 3]))).toMatchObject({
    success: false,
    code: 'invalid-document',
  });
  expect(prepareReviewGroupDocument(null as unknown as Uint8Array)).toMatchObject({
    success: false,
    code: 'invalid-input',
  });
});
