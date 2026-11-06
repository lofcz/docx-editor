import { expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart, canonicalOoxmlFingerprint } from '../index';
import { apply } from './tracked-edit-fixture';
import marked from './fixtures/run-origins/paragraph-role-model1-reader.json';
import unmarked from './fixtures/run-origins/paragraph-mark-model1-reader.json';

for (const [role, fixture] of [
  ['marked', marked],
  ['unmarked', unmarked],
] as const)
  test(`model 1 reader evidence preserves ${role} paragraph ownership and pending XML`, () => {
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
      // Legacy marks never own ordinary run proofs. Only explicit paragraph-source roles
      // activate the model-2 restoration contract; imported unmarked data stays opaque.
      expect(serializeOoxmlPart(decided)).toBe(
        role === 'marked' ? fixture[`${action}ExpectedXml`] : fixture[`${action}Xml`]
      );
    }
  });
