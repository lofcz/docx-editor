import { expect, test } from 'bun:test';
import { part, paragraphId, apply } from './tracked-edit-fixture.ts';
import { readOoxmlPart, canonicalOoxmlFingerprint, prepareTrackedRunOrigins } from '../index.ts';
const data = await Bun.file(
  new URL('./fixtures/run-origins/installed-reader-output.json', import.meta.url)
).json();
const input = await Bun.file(
  new URL('./fixtures/run-origins/mixed-reader-input.json', import.meta.url)
).json();
function read(xml: string) {
  const r = readOoxmlPart(xml, { name: input.name, contentType: input.contentType });
  if (!r.ok) throw Error(r.reason);
  return r.part;
}
for (const action of ['accept', 'reject'] as const)
  test(`an installed-reader ordinary edit survives new-reader ${action}`, () => {
    const current = read(data.editedXml);
    const source = part('<w:p><w:r><w:rPr/><w:t>AB</w:t></w:r></w:p>');
    const id = paragraphId(source);
    let golden =
      action === 'accept'
        ? apply(source, { op: 'insertText', paragraphId: id, offset: 1, text: 'X' })
        : source;
    golden = apply(golden, {
      op: 'insertText',
      paragraphId: id,
      offset: action === 'accept' ? 3 : 2,
      text: '!',
    });
    const decided = apply(current, {
      op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
    });
    expect(canonicalOoxmlFingerprint(decided)).toBe(canonicalOoxmlFingerprint(golden));
  });
test('a later proposal remains available after an installed-reader decision', () => {
  const current = read(data.acceptedXml);
  expect(() =>
    prepareTrackedRunOrigins(current, paragraphId(current), [{ start: 1, end: 1 }])
  ).not.toThrow();
});
