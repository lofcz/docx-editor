import { reviewGroupRevisionReservations } from './group-allocation';
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  collectRevisionSites,
  withCustomXmlNode,
  canonicalOoxmlFingerprint,
  findNode,
  withPart,
  contentControlContentNodeOf,
  type OoxmlNode,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import {
  presenceAttribution,
  presencePackage,
  presenceRecord,
  reopenPresence,
} from '../../test/support/block-presence';
import {
  decodeBlockPresenceRecord,
  encodeBlockPresenceRecord,
  type BlockPresenceRecord,
} from './block-presence-schema';
import { createBlockPresenceRecord, bindBlockPresenceRecord } from './block-presence-plan';
import { projectBlockPresence } from './block-presence-projection';
import { withBlockPresenceRecord } from './block-presence-store';
import { readPropertyReviewStore } from './property-record-store';
import { decodePropertyReviewRecord } from './property-record-schema';
import { readStoredPropertyProposals } from './stored-property-proposals';
const main = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
for (const remove of [false, true])
  test(`presence ${remove ? 'removal' : 'insertion'} owns a complete nested payload set and explicit absence through Word reopen`, () => {
    const pkg = presencePackage(),
      record = presenceRecord(undefined, pkg, remove);
    expect(record.blocks[0]!.before).toEqual({ present: remove });
    expect(record.blocks[0]!.after).toEqual({ present: !remove });
    expect(record.blocks[0]!.members.map((member) => member.payload.id)).toEqual([
      'outer',
      'inner',
    ]);
    const encoded = encodeBlockPresenceRecord(record),
      decoded = decodeBlockPresenceRecord(encoded);
    expect(decoded.status).toBe('supported');
    if (decoded.status !== 'supported') throw Error('Missing record');
    expect(Object.isFrozen(decoded.value.blocks[0]!.members[1]!.payload.extension)).toBe(true);
    const stored = withBlockPresenceRecord(pkg, decoded.value);
    if (!stored.success) throw Error(stored.error);
    const reopened = reopenPresence(stored.pkg);
    const source = readPropertyReviewStore(reopened);
    if (source.status !== 'ready') throw Error('Missing store');
    expect(source.records.get(record.id)?.raw).toBe(encoded);
    expect(bindBlockPresenceRecord(reopened, record).success).toBe(true);
    expect(withBlockPresenceRecord(stored.pkg, record)).toEqual({ success: true, pkg: stored.pkg });
    const hidden = projectBlockPresence(reopened, record, remove ? 'working' : 'committed');
    if (!hidden.success) throw Error(hidden.error);
    expect(canonicalOoxmlFingerprint(main(hidden.pkg))).toBe(
      canonicalOoxmlFingerprint(main(presencePackage(false)))
    );
    expect(buildBlockIndex(hidden.pkg).byId.size).toBe(0);
    expect(buildBlockIndex(hidden.pkg).diagnostics).toEqual([]);
    const visible = projectBlockPresence(reopened, record, remove ? 'committed' : 'working');
    if (!visible.success) throw Error(visible.error);
    expect(visible.pkg).toBe(reopened);
    expect(bindBlockPresenceRecord(reopened, record).success).toBe(true);
    expect(main(stored.pkg)).toBe(main(pkg));
    for (const [name, part] of reopened.parts)
      if (
        name !== reopened.mainDocumentPart &&
        name !== (readPayloadStore(reopened) as { part: { partName: string } }).part.partName
      )
        expect(hidden.pkg.parts.get(name)).toBe(part);
  });
test('legacy property readers preserve presence bytes and explicitly refuse unsupported version two', () => {
  const pkg = presencePackage(),
    record = presenceRecord(undefined, pkg),
    raw = encodeBlockPresenceRecord(record);
  expect(decodePropertyReviewRecord(raw)).toEqual({
    status: 'unsupported-version',
    id: record.id,
    version: 2,
    raw,
  });
  const stored = withBlockPresenceRecord(pkg, record);
  if (!stored.success) throw Error(stored.error);
  expect(reviewGroupRevisionReservations(stored.pkg)).toEqual([]);
  const before = canonicalOoxmlFingerprint(main(stored.pkg));
  const legacy = readStoredPropertyProposals(stored.pkg, {
    documentId: 'presence',
    epoch: 'one',
    revision: 0,
  });
  expect(legacy.success).toBe(false);
  if (!legacy.success) expect(legacy.code).toBe('unsupported-operation');
  expect(canonicalOoxmlFingerprint(main(stored.pkg))).toBe(before);
  const future = '{ "schemaVersion": 3, "id": "future", "unknown": ["Žák", null] }';
  expect(decodeBlockPresenceRecord(future)).toEqual({
    status: 'unsupported-version',
    id: 'future',
    version: 3,
    raw: future,
  });
});
test('presence ownership survives a fresh load without trusting saved canonical node IDs', () => {
  const before = presencePackage(),
    reopened = reopenPresence(before);
  const remap = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'textValue'
      ? { ...node, id: 'fresh:' + node.id }
      : ({ ...node, id: 'fresh:' + node.id, children: node.children.map(remap) } as OoxmlElement);
  const after = withPart(reopened, {
    ...main(reopened),
    root: remap(main(reopened).root) as OoxmlElement,
  });
  expect(buildBlockIndex(before).byId.get('outer')!.controlId).not.toBe(
    buildBlockIndex(after).byId.get('outer')!.controlId
  );
  const removal = createBlockPresenceRecord(
    before,
    after,
    [{ blockId: 'outer', before: { present: true }, after: { present: false } }],
    presenceAttribution
  );
  expect(removal.success).toBe(true);
});
for (const change of ['native', 'payload', 'label'] as const)
  test(`changed ${change} ownership refuses projection and storage before writes`, () => {
    const pkg = presencePackage(),
      record = presenceRecord(undefined, pkg),
      index = buildBlockIndex(pkg),
      child = index.byId.get('inner')!;
    let changed = pkg;
    if (change === 'native') {
      const part = main(pkg),
        root = findNode(part, child.controlId)!;
      if (root.kind === 'textValue') throw Error('Missing root');
      const paragraph = contentControlContentNodeOf(root)?.children.find(
        (node) => node.kind === 'paragraph'
      );
      if (!paragraph) throw Error('Missing fallback');
      const edit = applyTreeOp(part, {
        op: 'insertText',
        paragraphId: paragraph.id,
        offset: 0,
        text: 'Manual ',
      });
      if (!edit.ok) throw Error(edit.reason);
      changed = withPart(pkg, edit.part);
    } else {
      const written = withPayloadRecord(
        pkg,
        change === 'payload' ? { ...child.payload, attributes: { changed: true } } : child.payload,
        change === 'label' ? 'Manual label' : 'inner label'
      );
      if (!written.ok) throw Error(written.reason);
      changed = written.pkg;
    }
    const before = main(changed);
    expect(bindBlockPresenceRecord(changed, record).success).toBe(false);
    expect(projectBlockPresence(changed, record, 'committed').success).toBe(false);
    expect(withBlockPresenceRecord(changed, record).success).toBe(false);
    expect(main(changed)).toBe(before);
    expect(
      readPropertyReviewStore(changed).status === 'ready' &&
        (readPropertyReviewStore(changed) as { records: ReadonlyMap<string, unknown> }).records.size
    ).toBe(0);
  });
test('nested roots, reused identities, and immutable proposal rewrites refuse', () => {
  const before = presencePackage(false),
    pkg = presencePackage();
  const both = createBlockPresenceRecord(
    before,
    pkg,
    ['outer', 'inner'].map((blockId) => ({
      blockId,
      before: { present: false },
      after: { present: true },
    })),
    presenceAttribution
  );
  expect(both.success).toBe(false);
  expect(
    createBlockPresenceRecord(
      pkg,
      pkg,
      [{ blockId: 'outer', before: { present: false }, after: { present: true } }],
      presenceAttribution
    ).success
  ).toBe(false);
  const record = presenceRecord(before, pkg),
    saved = withBlockPresenceRecord(pkg, record);
  if (!saved.success) throw Error(saved.error);
  expect(withBlockPresenceRecord(saved.pkg, { ...record, author: 'Other' }).success).toBe(false);
});
test('schema guards reject ambiguous states, unbounded identities, invalid paths, and non-data input', () => {
  const record = presenceRecord(),
    block = record.blocks[0]!;
  const variants = [
    { ...record, blocks: [{ ...block, before: block.after }] },
    { ...record, blocks: [{ ...block, partName: '/word/../escape.xml' }] },
    { ...record, blocks: [{ ...block, nativeFingerprint: 'sha256:bad' }] },
    { ...record, blocks: [{ ...block, members: [...block.members, block.members[0]] }] },
    { ...record, predecessors: [record.id] },
    { ...record, date: '2026-02-30T00:00:00.000Z' },
    { ...record, blocks: [{ ...block, after: { present: true, extra: 'discard' } }] },
    { ...record, extension: { bad: undefined } },
    { ...record, blocks: [{ ...block, members: [{ label: 'No payload' }] }] },
  ];
  for (const variant of variants)
    expect(() => encodeBlockPresenceRecord(variant as BlockPresenceRecord)).toThrow();
  let called = false;
  const accessor = {
    ...record,
    get extension() {
      called = true;
      return {};
    },
  };
  expect(() => encodeBlockPresenceRecord(accessor)).toThrow();
  expect(called).toBe(false);
});

test('presence preflight keeps pending native changes outside its ownership', () => {
  const pkg = presencePackage(),
    part = main(pkg),
    owner = buildBlockIndex(pkg).byId.get('inner')!;
  const node = findNode(part, owner.controlId)!;
  const paragraph = contentControlContentNodeOf(node)!.children.find(
    (child) => child.kind === 'paragraph'
  )!;
  const changed = applyTreeOp(part, {
    op: 'insertText',
    paragraphId: paragraph.id,
    offset: 0,
    text: 'Foreign ',
    revision: { author: 'Other', date: presenceAttribution.date },
  });
  if (!changed.ok) throw Error(changed.reason);
  const canonical = withPart(pkg, changed.part),
    before = canonicalOoxmlFingerprint(changed.part);
  const planned = createBlockPresenceRecord(
    canonical,
    canonical,
    [{ blockId: 'outer', before: { present: true }, after: { present: false } }],
    presenceAttribution
  );
  expect(planned.success).toBe(false);
  if (!planned.success) expect(planned.code).toBe('unsupported-operation');
  expect(canonicalOoxmlFingerprint(main(canonical))).toBe(before);
  expect(collectRevisionSites(main(canonical))).toHaveLength(1);
});
test('nested pending presence, resolved proposal IDs, and retired block identities cannot be donated to new proposals', () => {
  const pkg = presencePackage(),
    record = presenceRecord(undefined, pkg),
    written = withBlockPresenceRecord(pkg, record);
  if (!written.success) throw Error(written.error);
  const child = createBlockPresenceRecord(
    written.pkg,
    written.pkg,
    [{ blockId: 'inner', before: { present: true }, after: { present: false } }],
    { ...presenceAttribution, id: 'child-removal' }
  );
  expect(child.success).toBe(false);
  if (!child.success) expect(child.code).toBe('unsupported-operation');
  const absent = projectBlockPresence(written.pkg, record, 'committed');
  if (!absent.success) throw Error(absent.error);
  expect(
    createBlockPresenceRecord(
      absent.pkg,
      pkg,
      [{ blockId: 'outer', before: { present: false }, after: { present: true } }],
      { ...presenceAttribution, id: 'reused-block' }
    ).success
  ).toBe(false);
  const store = readPropertyReviewStore(written.pkg);
  if (store.status !== 'ready') throw Error('Missing store');
  const retired = withCustomXmlNode(written.pkg, store.part.partName, {
    id: 'scio-review_' + record.id,
    label: 'removed',
    data: encodeBlockPresenceRecord(record),
  });
  expect(withBlockPresenceRecord(retired, record).success).toBe(false);
});
