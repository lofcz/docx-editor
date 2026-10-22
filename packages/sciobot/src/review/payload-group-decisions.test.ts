import { expect, test } from 'bun:test';
import {
  customXmlNodes,
  readOoxmlPackage,
  withCustomXmlNode,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  decisionFixture,
  decisionReceipt,
  decisionState,
  decisionStates,
  multiBlockProposal,
  written,
} from '../../test/support/payload-group-decisions';
import {
  decodePayloadGroupDecisionRecord,
  encodePayloadGroupDecisionRecord,
  PAYLOAD_GROUP_DECISION_LIMITS,
  type PayloadGroupDecisionRecord,
} from './payload-group-decision-schema';
import { decodeReviewGroupRecord } from './group-record-schema';
import {
  readReviewGroupStore,
  withInitializedReviewGroupStore,
  withPayloadGroupDecisionRecord,
  withReviewGroupRecord,
  withoutReviewGroupRecord,
} from './group-record-store';
import { readPayloadGroupDecisions } from './payload-group-decisions';
import { readPropertyReviewStore } from './property-record-store';
import { reviewGroupRevisionReservations } from './group-allocation';
import { readReviewGroups } from './groups';

function initialized() {
  return written(withInitializedReviewGroupStore(decisionFixture()));
}
function source(pkg: OoxmlPackage) {
  const value = readReviewGroupStore(pkg);
  if (value.status !== 'ready') throw new Error('Missing carrier');
  return value;
}
function reopened(pkg: OoxmlPackage) {
  const result = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!result.ok) throw new Error(result.reason);
  return result.package;
}

test('decision receipts own bounded immutable stable identities without canonical addresses', () => {
  const record = { ...decisionReceipt() },
    raw = encodePayloadGroupDecisionRecord(record),
    decoded = decodePayloadGroupDecisionRecord(raw);
  record.blockId = 'changed';
  expect(decoded).toEqual({ status: 'supported', value: decisionReceipt() });
  expect(decoded.status === 'supported' && Object.isFrozen(decoded.value)).toBe(true);
  expect(decodeReviewGroupRecord(raw)).toEqual({
    status: 'unsupported-version',
    id: 'rg_alice',
    version: 2,
    raw,
  });
  const invalid: unknown[] = [
    null,
    [],
    { ...decisionReceipt(), id: 'bad/id' },
    { ...decisionReceipt(), proposalId: '../proposal' },
    { ...decisionReceipt(), blockId: 'a'.repeat(128) },
    { ...decisionReceipt(), decision: 'pending' },
    { ...decisionReceipt(), schemaVersion: 1 },
    { ...decisionReceipt(), schemaVersion: 3 },
    { ...decisionReceipt(), nodeId: 'load-specific' },
    { ...decisionReceipt(), path: [0, 1] },
    { ...decisionReceipt(), proposalId: undefined },
    { ...decisionReceipt(), extension: { extra: true } },
  ];
  for (const value of invalid) {
    expect(() => encodePayloadGroupDecisionRecord(value as PayloadGroupDecisionRecord)).toThrow();
    if (value !== undefined)
      expect(decodePayloadGroupDecisionRecord(JSON.stringify(value)).status).toBe('invalid');
  }
  expect(decodePayloadGroupDecisionRecord('invalid')).toEqual({
    status: 'invalid',
    reason: 'invalid-json',
  });
  expect(
    decodePayloadGroupDecisionRecord('x'.repeat(PAYLOAD_GROUP_DECISION_LIMITS.maxBytes + 1))
  ).toEqual({ status: 'invalid', reason: 'byte-limit' });
  let calls = 0;
  expect(() =>
    encodePayloadGroupDecisionRecord({
      ...decisionReceipt(),
      get blockId() {
        calls++;
        return 'activity';
      },
    })
  ).toThrow('non-json-property');
  expect(calls).toBe(0);
  const cyclic: Record<string, unknown> = {};
  cyclic.self = cyclic;
  expect(() =>
    encodePayloadGroupDecisionRecord({ ...decisionReceipt(), blockId: cyclic } as never)
  ).toThrow('cyclic-payload');
  expect(
    decodePayloadGroupDecisionRecord(
      JSON.stringify({ ...decisionReceipt(), blockId: [[[[['deep']]]]] })
    )
  ).toEqual({ status: 'invalid', reason: 'depth-limit' });
  expect(
    decodePayloadGroupDecisionRecord(
      JSON.stringify({ ...decisionReceipt(), blockId: Array.from({ length: 33 }, () => 0) })
    )
  ).toEqual({ status: 'invalid', reason: 'node-limit' });
});

test('legacy proposals use the existing initialized carrier without changing property JSON or native XML', () => {
  const legacy = decisionFixture(),
    before = writeOoxmlPackage(legacy);
  expect(withPayloadGroupDecisionRecord(legacy, decisionReceipt())).toMatchObject({
    ok: false,
    pkg: legacy,
    reason: 'review-group-store-initialization-required',
  });
  expect(readPayloadGroupDecisions(legacy)).toMatchObject({ status: 'ready' });
  expect(writeOoxmlPackage(legacy)).toEqual(before);
  const baseline = written(withInitializedReviewGroupStore(legacy)),
    first = written(withPayloadGroupDecisionRecord(baseline, decisionReceipt())),
    second = written(
      withPayloadGroupDecisionRecord(first, decisionReceipt('rg_bob', 'phase', 'reject'))
    ),
    propertyStore = readPropertyReviewStore(legacy);
  if (propertyStore.status !== 'ready') throw new Error('Missing original proposal');
  expect(second.parts.get(propertyStore.part.partName)).toBe(
    legacy.parts.get(propertyStore.part.partName)
  );
  expect(second.parts.get(legacy.mainDocumentPart)).toBe(legacy.parts.get(legacy.mainDocumentPart));
  expect(written(withInitializedReviewGroupStore(second))).toBe(second);
  expect(decisionStates(second)).toEqual([
    ['activity', 'accepted', ['rg_alice']],
    ['phase', 'rejected', ['rg_bob']],
  ]);
  expect(decisionStates(reopened(second))).toEqual(decisionStates(second));
  const savedStore = readPropertyReviewStore(reopened(second));
  expect(savedStore.status === 'ready' && savedStore.records.get(multiBlockProposal.id)!.raw).toBe(
    propertyStore.records.get(multiBlockProposal.id)!.raw
  );
  expect(reviewGroupRevisionReservations(second)).toEqual([]);
  // R15 consumes persisted dispositions without treating receipts as native membership.
  expect(readReviewGroups(second)).toEqual({ success: true, groups: [] });
});

test('independent receipt identities retain equal and conflicting same-group decisions without a winner', () => {
  const first = written(withPayloadGroupDecisionRecord(initialized(), decisionReceipt())),
    same = written(withPayloadGroupDecisionRecord(first, decisionReceipt('rg_same'))),
    conflict = written(
      withPayloadGroupDecisionRecord(same, decisionReceipt('rg_opposite', 'activity', 'reject'))
    );
  expect(decisionState(same)).toMatchObject({
    status: 'accepted',
    receiptIds: ['rg_alice', 'rg_same'],
  });
  expect(decisionState(conflict)).toMatchObject({
    status: 'conflict',
    receiptIds: ['rg_alice', 'rg_opposite', 'rg_same'],
  });
  expect(decisionState(reopened(conflict))).toEqual(decisionState(conflict));
  expect(withPayloadGroupDecisionRecord(first, decisionReceipt())).toEqual({
    ok: true,
    pkg: first,
  });
  expect(withPayloadGroupDecisionRecord(first, decisionReceipt('rg_alice', 'phase'))).toMatchObject(
    {
      ok: false,
      pkg: first,
      reason: 'review-group-record-is-immutable',
    }
  );
  expect(withoutReviewGroupRecord(first, 'rg_alice')).toMatchObject({
    ok: false,
    pkg: first,
    reason: 'review-group-record-is-immutable',
  });
  const native = {
    schemaVersion: 1 as const,
    id: 'rg_native',
    members: [{ partName: '/word/document.xml', revision: { id: '123', author: 'Teacher' } }],
  };
  const combined = written(withReviewGroupRecord(conflict, native));
  expect(reviewGroupRevisionReservations(combined)).toEqual([
    { partName: '/word/document.xml', ids: ['123'] },
  ]);
  expect(withPayloadGroupDecisionRecord(combined, decisionReceipt('rg_native'))).toMatchObject({
    ok: false,
  });
  expect(withReviewGroupRecord(combined, { ...native, id: 'rg_alice' })).toMatchObject({
    ok: false,
  });
  const removed = written(withoutReviewGroupRecord(combined, native.id));
  expect(withPayloadGroupDecisionRecord(removed, decisionReceipt('rg_native'))).toMatchObject({
    ok: false,
    reason: 'review-group-id-reserved',
  });
});

test('future data stays exact and refuses decisions while malformed known receipts fail before writes', () => {
  const pkg = initialized(),
    descriptor = source(pkg).part,
    raw = '{ "schemaVersion": 3, "id": "rg_future", "opaque": [1, " keep "] }',
    future = withCustomXmlNode(pkg, descriptor.partName, { id: 'rg_future', label: '', data: raw }),
    candidate = written(withPayloadGroupDecisionRecord(future, decisionReceipt()));
  expect(source(reopened(candidate)).records.get('rg_future')!.raw).toBe(raw);
  expect(withPayloadGroupDecisionRecord(candidate, decisionReceipt('rg_future'))).toMatchObject({
    ok: false,
    pkg: candidate,
    reason: 'unsupported-review-group-version',
  });
  expect(readPayloadGroupDecisions(candidate)).toEqual({
    status: 'invalid',
    reason: 'unsupported-review-group-version',
  });
  for (const node of [
    { id: 'rg_alice', label: 'removed', data: encodePayloadGroupDecisionRecord(decisionReceipt()) },
    { id: 'rg_other', label: '', data: encodePayloadGroupDecisionRecord(decisionReceipt()) },
    {
      id: 'rg_alice',
      label: '',
      data: JSON.stringify({ ...decisionReceipt(), canonicalId: 'unstable' }),
    },
  ]) {
    const malformed = withCustomXmlNode(pkg, descriptor.partName, node),
      before = writeOoxmlPackage(malformed);
    expect(readReviewGroupStore(malformed).status).toBe('invalid');
    expect(withPayloadGroupDecisionRecord(malformed, decisionReceipt('rg_new'))).toMatchObject({
      ok: false,
      pkg: malformed,
    });
    expect(writeOoxmlPackage(malformed)).toEqual(before);
  }
  expect(
    customXmlNodes(candidate, descriptor.partName).find((node) => node.id === 'rg_future')!.data
  ).toBe(raw);
});
