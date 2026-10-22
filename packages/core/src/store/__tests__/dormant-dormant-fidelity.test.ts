import { test, expect } from 'bun:test';
import {
  readOoxmlPart,
  serializeOoxmlPart,
  canonicalOoxmlFingerprint,
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
} from '../index.ts';
import { paragraphId, apply, ADA } from './tracked-edit-fixture.ts';
const data = await Bun.file(
  new URL('./fixtures/run-origins/installed-reader-output.json', import.meta.url)
).json();
const input = await Bun.file(
  new URL('./fixtures/run-origins/mixed-reader-input.json', import.meta.url)
).json();
const fp = canonicalOoxmlFingerprint;
function read(xml: string) {
  const r = readOoxmlPart(xml, { name: input.name, contentType: input.contentType });
  if (!r.ok) throw Error(r.reason);
  return r.part;
}
for (const action of ['accept', 'reject'] as const)
  for (const saved of [false, true])
    test(`later proposal preserves dormant source attributes ${action}/${saved}`, () => {
      const source = read(data.acceptedXml),
        id = paragraphId(source),
        before = fp(source);
      const prepared = prepareTrackedRunOrigins(source, id, [{ start: 1, end: 1 }]);
      let pending = apply(prepared.part, {
        op: 'insertText',
        paragraphId: id,
        offset: 1,
        text: 'Y',
        revision: { ...ADA, date: '2026-10-22T12:00:00Z' },
      });
      pending = sealTrackedRunOrigins(pending, id, prepared.scope!, source);
      expect(() => prepareTrackedRunOrigins(pending, id, [{ start: 0, end: 1 }])).toThrow(
        'active or opaque'
      );
      if (saved) pending = read(serializeOoxmlPart(pending));
      const decided = apply(pending, {
        op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
      });
      const golden =
        action === 'accept'
          ? apply(source, { op: 'insertText', paragraphId: id, offset: 1, text: 'Y' })
          : source;
      expect(fp(decided)).toBe(fp(golden));
      expect(fp(source)).toBe(before);
    });
