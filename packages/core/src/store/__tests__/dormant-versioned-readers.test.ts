import { expect, test } from 'bun:test';
import {
  readOoxmlPart,
  serializeOoxmlPart,
  canonicalOoxmlFingerprint,
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  type OoxmlPart,
} from '../index.ts';
import { apply, paragraphId, ADA } from './tracked-edit-fixture.ts';
const fp = canonicalOoxmlFingerprint;
for (const fixture of ['dynamic', 'layered', 'current'])
  for (const priorAction of ['accept', 'reject'] as const) {
    const input = await Bun.file(
      new URL(`./fixtures/run-origins/${fixture}-mixed-reader-input.json`, import.meta.url)
    ).json();
    const data = await Bun.file(
      new URL(`./fixtures/run-origins/${fixture}-installed-reader-output.json`, import.meta.url)
    ).json();
    const read = (xml: string): OoxmlPart => {
      const r = readOoxmlPart(xml, { name: input.name, contentType: input.contentType });
      if (!r.ok) throw Error(r.reason);
      return r.part;
    };
    for (const saved of [false, true])
      for (const action of ['accept', 'reject'] as const)
        test(`later proposal after installed-reader ${fixture}/${priorAction} decisions ${saved}/${action}`, () => {
          const source = read(priorAction === 'accept' ? data.acceptedXml : data.rejectedXml),
            before = fp(source),
            id = paragraphId(source);
          const prepared = prepareTrackedRunOrigins(source, id, [{ start: 1, end: 1 }]);
          let pending = sealTrackedRunOrigins(
            apply(prepared.part, {
              op: 'insertText',
              paragraphId: id,
              offset: 1,
              text: 'Z',
              revision: { ...ADA, date: '2026-10-22T13:00:00Z' },
            }),
            id,
            prepared.scope!,
            source
          );
          if (saved) pending = read(serializeOoxmlPart(pending));
          const actual = apply(pending, {
            op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
          });
          const expected =
            action === 'accept'
              ? apply(source, { op: 'insertText', paragraphId: id, offset: 1, text: 'Z' })
              : source;
          expect(fp(actual)).toBe(fp(expected));
          expect(fp(source)).toBe(before);
        });
  }
