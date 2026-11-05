import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart, canonicalOoxmlFingerprint } from '../index';
import { apply } from './tracked-edit-fixture';
import fixture from './fixtures/run-origins/paragraph-mark-model1-reader.json';

test('model 1 reader evidence retains pending XML but leaves stale paragraph source proofs after decisions', () => {
  expect(fixture.reader).toContain('fb6f6405b');
  expect(fixture.readerSha256).toMatch(/^[0-9a-f]{64}$/);
  expect(fixture.oldPendingXml).toBe(fixture.pendingXml);
  for (const action of ['accept', 'reject'] as const) {
    expect(fixture[`${action}Xml`]).not.toBe(fixture[`${action}ExpectedXml`]);
    expect(fixture[`${action}Xml`]).toContain('tracked-run-origins:1/');
    const read = readOoxmlPart(fixture.pendingXml, {
      name: fixture.name,
      contentType: fixture.contentType,
    });
    if (!read.ok) throw Error(read.reason);
    const before = canonicalOoxmlFingerprint(read.part);
    expect(serializeOoxmlPart(read.part)).toBe(fixture.pendingXml);
    expect(canonicalOoxmlFingerprint(read.part)).toBe(before);
    const decided = apply(read.part, {
      op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
    });
    expect(serializeOoxmlPart(decided)).toBe(fixture[`${action}ExpectedXml`]);
  }
});
