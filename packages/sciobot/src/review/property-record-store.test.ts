import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  customXmlNodes,
  escapeXml,
  findNode,
  headlessViewOfStore,
  readOoxmlPackage,
  readOoxmlPart,
  TreePackageStore,
  withCustomXmlDataPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  packageWithPropertyRecord,
  propertyActivity,
  propertyRecord,
  propertyRecordFixture,
  propertyRecordRevision,
} from '../../test/support/property-record';
import { sourceControl, sourceParagraph } from '../../test/support/target-source-fixture';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import type { JsonObject } from '../data/json';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import {
  decodePropertyReviewRecord,
  encodePropertyReviewRecord,
  type PropertyReviewRecord,
} from './property-record-schema';
import {
  PROPERTY_REVIEW_NAMESPACE,
  PROPERTY_REVIEW_ROOT,
  readPropertyReviewStore,
  withPropertyReviewRecord,
  withoutPropertyReviewRecord,
} from './property-record-store';
import { readStoredPropertyProposals } from './stored-property-proposals';

function ready(pkg: OoxmlPackage) {
  const read = readPropertyReviewStore(pkg);
  if (read.status !== 'ready') throw new Error('Property store is not ready');
  return read;
}
function replaceStore(pkg: OoxmlPackage, body: string, root = PROPERTY_REVIEW_ROOT) {
  const current = ready(pkg),
    part = pkg.parts.get(current.part.partName)!,
    loaded = readOoxmlPart(`<${root} xmlns="${PROPERTY_REVIEW_NAMESPACE}">${body}</${root}>`, {
      name: part.name,
      contentType: part.contentType,
    });
  if (!loaded.ok) throw new Error(loaded.reason);
  return withPart(pkg, loaded.part);
}
const node = (id: string, raw: string, extra = '') =>
  `<node id="scio-review_${id}"><label/><data>${escapeXml(raw)}</data>${extra}</node>`;
function reopened(pkg: OoxmlPackage) {
  const loaded = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.package;
}

test('records preserve explicit property presence, native addresses, attribution, and extension data through native save/reopen', () => {
  const pkg = propertyRecordFixture(),
    original = writeOoxmlPackage(pkg),
    input: PropertyReviewRecord = {
      ...propertyRecord(),
      blocks: [
        {
          blockId: 'activity',
          type: 'lesson_activity',
          changes: [
            {
              name: 'duration',
              before: { present: true, value: 10 },
              after: { present: true, value: 15 },
            },
            { name: 'name', before: { present: true, value: null }, after: { present: false } },
            {
              name: 'resource',
              before: { present: false },
              after: {
                present: true,
                value: { zero: 0, enabled: false, nullable: null, nested: ['Žák 😀'] },
              },
            },
          ],
        },
      ],
    },
    current = packageWithPropertyRecord(pkg, input),
    saved = reopened(current);
  for (const source of [current, saved]) {
    expect(ready(source).records.get(input.id)!.decoded).toEqual({
      status: 'supported',
      value: input,
    });
    expect(propertyActivity(source).payload).toEqual(propertyActivity(pkg).payload);
    expect(buildBlockIndex(source).byId.get('phase')!.payload.attributes).toEqual({
      phaseType: 'evocation',
    });
    const ownStore = ready(source);
    expect(
      customXmlNodes(source, ownStore.part.partName)
        .filter((entry) => entry.id.startsWith('scio-review_'))
        .map((entry) => entry.label)
    ).toEqual(['']);
  }
  expect(current.parts.get(pkg.mainDocumentPart)).toBe(pkg.parts.get(pkg.mainDocumentPart));
  expect(readPayloadStore(current)).toEqual(readPayloadStore(pkg));
  expect(writeOoxmlPackage(pkg)).toEqual(original);
});

test('encoded records own and freeze their data while prototype-shaped JSON stays inert', () => {
  const record = {
      ...propertyRecord(),
      extension: JSON.parse(
        '{"__proto__":{"polluted":true},"constructor":{"prototype":{"keep":1}}}'
      ) as JsonObject,
    },
    raw = encodePropertyReviewRecord(record),
    decoded = decodePropertyReviewRecord(raw);
  if (decoded.status !== 'supported') throw new Error('Invalid record');
  expect(Object.isFrozen(decoded.value)).toBe(true);
  expect(Object.isFrozen(decoded.value.blocks[0]!.changes[0]!.before)).toBe(true);
  expect(Object.isFrozen(decoded.value.extension)).toBe(true);
  expect(Object.hasOwn(decoded.value.extension!, '__proto__')).toBe(true);
  expect(({} as { polluted?: boolean }).polluted).toBeUndefined();
  (record.extension as Record<string, unknown>).__proto__ = 'changed';
  expect(decoded.value.extension!.__proto__).toEqual({ polluted: true });
});

test('record gates refuse ambiguous, malformed, unsafe, and unsupported input without invoking getters', () => {
  const invalid: unknown[] = [
    { ...propertyRecord(), schemaVersion: 0 },
    { ...propertyRecord(), id: 'bad/id' },
    { ...propertyRecord(), author: ' ' },
    { ...propertyRecord(), author: '\uD800' },
    { ...propertyRecord(), date: '2026-02-30T09:00:00.000Z' },
    { ...propertyRecord(), date: '2026-10-19T09:00:00Z' },
    { ...propertyRecord(), blocks: [] },
    { ...propertyRecord(), blocks: [propertyRecord().blocks[0], propertyRecord().blocks[0]] },
    { ...propertyRecord(), blocks: [{ ...propertyRecord().blocks[0], changes: [] }] },
    {
      ...propertyRecord(),
      blocks: [
        {
          ...propertyRecord().blocks[0],
          changes: [propertyRecord().blocks[0]!.changes[0], propertyRecord().blocks[0]!.changes[0]],
        },
      ],
    },
    {
      ...propertyRecord(),
      blocks: [
        {
          blockId: 'activity',
          type: 'lesson_activity',
          changes: [
            {
              name: 'duration',
              before: { present: false, value: 10 },
              after: { present: true, value: 15 },
            },
          ],
        },
      ],
    },
    {
      ...propertyRecord(),
      blocks: [
        {
          blockId: 'activity',
          type: 'lesson_activity',
          changes: [{ name: 'duration', before: { present: true }, after: { present: false } }],
        },
      ],
    },
    {
      ...propertyRecord(),
      nativeRevisions: [{ partName: '/word/../document.xml', id: '1', author: 'ScioBot' }],
    },
    {
      ...propertyRecord(),
      nativeRevisions: [propertyRecord().nativeRevisions[0], propertyRecord().nativeRevisions[0]],
    },
    { ...propertyRecord(), predecessors: [propertyRecord().id] },
    { ...propertyRecord(), predecessors: ['one', 'one'] },
    { ...propertyRecord(), extension: [] },
    { ...propertyRecord(), unknown: 1 },
    { ...propertyRecord(), extension: { huge: 'x'.repeat(2 * 1024 * 1024) } },
  ];
  for (const value of invalid)
    expect(() => encodePropertyReviewRecord(value as PropertyReviewRecord)).toThrow();
  expect(decodePropertyReviewRecord('invalid')).toEqual({
    status: 'invalid',
    reason: 'invalid-json',
  });
  expect(decodePropertyReviewRecord('x'.repeat(2 * 1024 * 1024 + 1))).toEqual({
    status: 'invalid',
    reason: 'byte-limit',
  });
  let calls = 0;
  const getter = {
    ...propertyRecord(),
    get author(): string {
      calls++;
      throw new Error('executed');
    },
  };
  expect(() => encodePropertyReviewRecord(getter)).toThrow('non-json-property');
  expect(calls).toBe(0);
  const cycle: Record<string, unknown> = {};
  cycle.self = cycle;
  expect(() =>
    encodePropertyReviewRecord({ ...propertyRecord(), extension: cycle as JsonObject })
  ).toThrow('cyclic-payload');
});

test('store updates retain native XML identities, foreign metadata, independent records, and original inputs', () => {
  const pkg = packageWithPropertyRecord(propertyRecordFixture()),
    record = propertyRecord(),
    raw = encodePropertyReviewRecord(record),
    imported = replaceStore(
      pkg,
      `<node id="scio-review_${record.id}" xmlns:keep="urn:keep" keep:flag="node"><label keep:flag="label"/><data keep:flag="data">${escapeXml(raw)}</data><keep:extra>Preserve</keep:extra></node>`
    ),
    part = imported.parts.get(ready(imported).part.partName)!,
    originalNode = part.root.children.find(
      (child) => child.kind !== 'textValue' && child.localName === 'node'
    )!,
    previous = writeOoxmlPackage(imported),
    updated = packageWithPropertyRecord(imported, { ...record, extension: { updated: true } }),
    withPeer = packageWithPropertyRecord(updated, propertyRecord('peer'));
  expect(withPropertyReviewRecord(withPeer, propertyRecord('peer'))).toEqual({
    ok: true,
    pkg: withPeer,
  });
  expect(ready(withPeer).records.size).toBe(2);
  const current = findNode(withPeer.parts.get(part.name)!, originalNode.id)!;
  expect(current.id).toBe(originalNode.id);
  if (current.kind === 'textValue' || originalNode.kind === 'textValue')
    throw new Error('Missing record node');
  expect(current.attributes).toEqual(originalNode.attributes);
  expect(
    current.children.find(
      (child) => child.kind !== 'textValue' && child.namespaceUri === 'urn:keep'
    )
  ).toBe(originalNode.children.at(-1));
  expect(withPeer.parts.get(pkg.mainDocumentPart)).toBe(imported.parts.get(pkg.mainDocumentPart));
  expect(writeOoxmlPackage(imported)).toEqual(previous);
  const result = withoutPropertyReviewRecord(withPeer, record.id);
  if (!result.ok) throw new Error(result.reason);
  expect([...ready(result.pkg).records.keys()]).toEqual(['peer']);
  const empty = withoutPropertyReviewRecord(result.pkg, 'peer');
  if (!empty.ok) throw new Error(empty.reason);
  expect(ready(empty.pkg).records.size).toBe(0);
  expect([...ready(empty.pkg).removedIds].sort()).toEqual(['peer', record.id]);
  expect(withPropertyReviewRecord(empty.pkg, record)).toMatchObject({
    ok: false,
    reason: 'removed-property-record-id',
  });
  expect(withoutPropertyReviewRecord(empty.pkg, 'missing')).toEqual({ ok: true, pkg: empty.pkg });
});

test('future records retain exact bytes and cannot be overwritten, removed, or silently read as known proposals', async () => {
  const pkg = packageWithPropertyRecord(propertyRecordFixture()),
    raw = '{ "schemaVersion": 2, "id": "future", "unknown": ["preserve"] }',
    future = replaceStore(pkg, node('future', raw)),
    before = writeOoxmlPackage(future);
  expect(ready(future).records.get('future')!.decoded).toEqual({
    status: 'unsupported-version',
    id: 'future',
    version: 2,
    raw,
  });
  expect(withPropertyReviewRecord(future, propertyRecord('future'))).toMatchObject({
    ok: false,
    reason: 'unsupported-property-record-version',
    pkg: future,
  });
  expect(withoutPropertyReviewRecord(future, 'future')).toMatchObject({
    ok: false,
    reason: 'unsupported-property-record-version',
    pkg: future,
  });
  const next = packageWithPropertyRecord(future, propertyRecord('known'));
  expect(ready(reopened(next)).records.get('future')!.raw).toBe(raw);
  expect(readStoredPropertyProposals(next, propertyRecordRevision)).toMatchObject({
    success: false,
    code: 'unsupported-operation',
  });
  const native = headlessViewOfStore(
    new TreePackageStore(future, future.parts.get(future.mainDocumentPart)!)
  );
  expect(await readWorkingView(native, propertyRecordRevision)).toMatchObject({
    success: false,
    code: 'unsupported-operation',
  });
  expect(writeOoxmlPackage(future)).toEqual(before);
});

test('duplicate stores, missing/duplicate identities, wrong roots, ambiguous fields, and count overflow refuse atomically', () => {
  const pkg = packageWithPropertyRecord(propertyRecordFixture()),
    raw = encodePropertyReviewRecord(propertyRecord()),
    id = propertyRecord().id;
  for (const invalid of [
    replaceStore(pkg, node(id, raw) + node(id, raw)),
    replaceStore(pkg, '<node><label/><data>{}</data></node>'),
    replaceStore(pkg, node('mismatch', raw)),
    replaceStore(pkg, node(id, raw, '<data>duplicate</data>')),
    replaceStore(
      pkg,
      `<node id="scio-review_${id}"><label>unexpected</label><data>${escapeXml(raw)}</data></node>`
    ),
    replaceStore(
      pkg,
      `<node id="scio-review_${id}"><label/><data><nested>${escapeXml(raw)}</nested></data></node>`
    ),
    replaceStore(pkg, '', 'wrongRoot'),
    replaceStore(pkg, Array.from({ length: 1025 }, (_, index) => node('p' + index, raw)).join('')),
  ]) {
    const bytes = writeOoxmlPackage(invalid);
    expect(readPropertyReviewStore(invalid).status).toBe('invalid');
    expect(withPropertyReviewRecord(invalid, propertyRecord()).ok).toBe(false);
    expect(withoutPropertyReviewRecord(invalid, id).ok).toBe(false);
    expect(writeOoxmlPackage(invalid)).toEqual(bytes);
  }
  const other = withCustomXmlDataPart(
    pkg,
    pkg.mainDocumentPart,
    PROPERTY_REVIEW_NAMESPACE,
    'otherRoot'
  );
  expect(other.pkg).toBe(pkg);
});

test('stable block IDs bind to fresh owners after source insertion and native reopen', () => {
  const pkg = packageWithPropertyRecord(propertyRecordFixture()),
    oldId = propertyActivity(pkg).controlId,
    zip = unzipSync(writeOoxmlPackage(pkg));
  zip['word/document.xml'] = strToU8(
    strFromU8(zip['word/document.xml']!).replace('<w:body>', '<w:body>' + sourceParagraph('Before'))
  );
  const loaded = readOoxmlPackage(zipSync(zip));
  if (!loaded.ok) throw new Error(loaded.reason);
  const current = loaded.package,
    owner = propertyActivity(current),
    result = readStoredPropertyProposals(current, { ...propertyRecordRevision, epoch: 'reload' });
  expect(owner.controlId).not.toBe(oldId);
  if (!result.success) throw new Error(result.error);
  expect(result.pending[0]!.plan.operations[0]).toMatchObject({
    blockId: 'activity',
    controlId: owner.controlId,
    partName: owner.partName,
    path: ['attributes', 'duration'],
  });
  expect(result.pending[0]!.plan.revision.epoch).toBe('reload');
  expect(JSON.stringify(ready(current).records.get(propertyRecord().id)!.decoded)).not.toContain(
    owner.controlId
  );
});

test('predecessors preserve value order independently of IDs, dates, and XML record order', async () => {
  const first = propertyRecord('z-first'),
    next: PropertyReviewRecord = {
      ...propertyRecord('a-second'),
      predecessors: [first.id],
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
    },
    pkg = packageWithPropertyRecord(
      packageWithPropertyRecord(propertyRecordFixture(), next),
      first
    ),
    source = headlessViewOfStore(new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!)),
    proposals = readStoredPropertyProposals(pkg, propertyRecordRevision);
  if (!proposals.success) throw new Error(proposals.error);
  expect(proposals.pending.map((item) => item.id)).toEqual(['z-first', 'a-second']);
  const [committed, working] = await Promise.all([
    readCommittedView(source, propertyRecordRevision),
    readWorkingView(source, propertyRecordRevision),
  ]);
  if (!committed.success || !working.success) throw new Error('Property views failed');
  expect(committed.view.projection.text).toContain('duration="10"');
  expect(working.view.projection.text).toContain('duration="20"');
  expect(propertyActivity(source.currentPackage()).payload.attributes.duration).toBe(10);
  const cycle = packageWithPropertyRecord(pkg, { ...first, predecessors: [next.id] });
  expect(readStoredPropertyProposals(cycle, propertyRecordRevision)).toMatchObject({
    success: false,
    code: 'invalid-document',
  });
  const concurrent = packageWithPropertyRecord(pkg, { ...next, predecessors: [] });
  expect(readStoredPropertyProposals(concurrent, propertyRecordRevision)).toMatchObject({
    success: false,
    code: 'ambiguous-target',
  });
});

test('an absent application store cannot acquire review records without canonical property owners', () => {
  const created = readOoxmlPackage(writeOoxmlPackage(propertyRecordFixture()));
  if (!created.ok) throw new Error(created.reason);
  const zip = unzipSync(writeOoxmlPackage(created.package));
  zip['word/_rels/document.xml.rels'] = strToU8(
    strFromU8(zip['word/_rels/document.xml.rels']!).replace(
      /<Relationship\b[^>]*Type="[^"]*\/customXml"[^>]*\/>/g,
      ''
    )
  );
  const absent = readOoxmlPackage(zipSync(zip));
  if (!absent.ok) throw new Error(absent.reason);
  expect(readPropertyReviewStore(absent.package)).toEqual({ status: 'absent' });
  expect(withPropertyReviewRecord(absent.package, propertyRecord())).toMatchObject({
    ok: false,
    pkg: absent.package,
    reason: 'application-payload-store-required',
  });
});

test('duplicate application store relationships refuse while empty native undo wrappers remain inert', () => {
  const pkg = packageWithPropertyRecord(propertyRecordFixture()),
    descriptor = ready(pkg).part,
    empty = replaceStore(
      pkg,
      node(propertyRecord().id, encodePropertyReviewRecord(propertyRecord())) + '<node/>'
    );
  expect(ready(empty).records.size).toBe(1);
  const zip = unzipSync(writeOoxmlPackage(pkg)),
    target = '..' + descriptor.partName;
  zip['word/_rels/document.xml.rels'] = strToU8(
    strFromU8(zip['word/_rels/document.xml.rels']!).replace(
      '</Relationships>',
      `<Relationship Id="duplicateStore" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="${target}"/></Relationships>`
    )
  );
  const loaded = readOoxmlPackage(zipSync(zip));
  if (!loaded.ok) throw new Error(loaded.reason);
  expect(readPropertyReviewStore(loaded.package)).toEqual({
    status: 'invalid',
    reason: 'duplicate-property-store',
  });
});

test('missing, duplicate, deleted, or type-changed native block owners never recreate a block or return a partial proposal', () => {
  const pkg = packageWithPropertyRecord(propertyRecordFixture());
  for (const change of [
    (xml: string) => xml.replace('scio:1:activity', 'foreign:activity'),
    (xml: string) =>
      xml.replace(
        '</w:body>',
        sourceControl('activity', sourceParagraph('Duplicate')) + '</w:body>'
      ),
  ]) {
    const bytes = unzipSync(writeOoxmlPackage(pkg));
    bytes['word/document.xml'] = strToU8(change(strFromU8(bytes['word/document.xml']!)));
    const loaded = readOoxmlPackage(zipSync(bytes));
    if (!loaded.ok) throw new Error(loaded.reason);
    const before = writeOoxmlPackage(loaded.package),
      result = readStoredPropertyProposals(loaded.package, propertyRecordRevision);
    expect(result.success).toBe(false);
    expect('pending' in result).toBe(false);
    expect(writeOoxmlPackage(loaded.package)).toEqual(before);
  }
  const bad = packageWithPropertyRecord(pkg, {
    ...propertyRecord(),
    blocks: [{ ...propertyRecord().blocks[0]!, type: 'lesson_title' }],
  });
  expect(readStoredPropertyProposals(bad, propertyRecordRevision)).toMatchObject({
    success: false,
    code: 'invalid-document',
  });
  expect(
    contentControlContentNodeOf(
      findNode(pkg.parts.get(pkg.mainDocumentPart)!, propertyActivity(pkg).controlId)!
    )
  ).not.toBeNull();
});

test('automatic review reads use stored properties and retain native paragraph targets without changing package state', async () => {
  const original = propertyRecordFixture(),
    pkg = packageWithPropertyRecord(original),
    native = headlessViewOfStore(new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!)),
    before = writeOoxmlPackage(pkg),
    fingerprint = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
  const [committed, working] = await Promise.all([
    readCommittedView(native, propertyRecordRevision),
    readWorkingView(native, propertyRecordRevision),
  ]);
  if (!committed.success) throw new Error(committed.error);
  if (!working.success) throw new Error(working.error);
  expect(committed.view.projection.text).toContain('duration="10"');
  expect(working.view.projection.text).toContain('duration="15"');
  for (const result of [committed, working]) {
    expect(result.view.projection.text).toContain('**Body**');
    expect(result.view.projection.text).toContain('**Outside**');
    expect(result.view.projection.blocks.map((item) => item.id)).toEqual(['phase', 'activity']);
  }
  expect(native.packageRevision()).toBe(0);
  expect(writeOoxmlPackage(native.currentPackage())).toEqual(before);
  expect(canonicalOoxmlFingerprint(native.part())).toBe(fingerprint);
  expect(readPayloadStore(pkg)).toEqual(readPayloadStore(original));
  const loaded = reopened(pkg),
    source = headlessViewOfStore(
      new TreePackageStore(loaded, loaded.parts.get(loaded.mainDocumentPart)!)
    ),
    after = await readWorkingView(source, { ...propertyRecordRevision, epoch: 'new-load' });
  if (!after.success) throw new Error(after.error);
  expect(after.view.projection.text).toBe(working.view.projection.text);
});
