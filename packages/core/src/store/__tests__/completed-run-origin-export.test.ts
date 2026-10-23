import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, readOoxmlPart, serializeOoxmlPart } from '../index.ts';
import { serializeOoxmlPartSteps } from '../package/ooxml-serialize-steps.ts';
import { projectCompletedRunOrigins } from '../store/completed-run-origin-export.ts';
import { part } from './tracked-edit-fixture.ts';

const scope = '12345678-1234-4234-8234-123456789abc';
const other = '87654321-1234-4234-8234-123456789abc';
const source = '12345678-1234-8234-a234-123456789abc';
const ns = `urn:docx-editor:tracked-run-origins:1/${scope}`;
const otherNs = `urn:docx-editor:tracked-run-origins:1/${other}`;
const fragment = (text: string, resolved = true) =>
  `<w:r xmlns:o="${ns}" o:origin="${scope}/${source}" o:identity="source"${resolved ? ' o:resolved="1"' : ''}><w:t o:textOrigin="${scope}/${source}" o:identity="source" o:space="absent"${resolved ? ' o:resolved="1"' : ''}>${text}</w:t></w:r>`;
const completed = `<w:p>${fragment('A')}${fragment('B')}</w:p>`;
function native(xml: string) {
  const read = readOoxmlPart(xml, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!read.ok) throw Error(read.reason);
  return canonicalOoxmlFingerprint(read.part);
}

test('completed export joins only proven source fragments, matches stepped output, and leaves the source unchanged', () => {
  const input = part(completed + '<w:p><w:r><w:t>C</w:t></w:r><w:r><w:t>D</w:t></w:r></w:p>');
  const before = canonicalOoxmlFingerprint(input);
  const xml = serializeOoxmlPart(input);
  expect(native(xml)).toBe(
    canonicalOoxmlFingerprint(
      part(
        '<w:p><w:r><w:t>AB</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r><w:r><w:t>D</w:t></w:r></w:p>'
      )
    )
  );
  const steps = serializeOoxmlPartSteps(structuredClone(input));
  let next = steps.next();
  while (!next.done) next = steps.next();
  expect(new TextDecoder().decode(next.value)).toBe(xml);
  expect(canonicalOoxmlFingerprint(input)).toBe(before);
  expect(projectCompletedRunOrigins(input).root).toBe(projectCompletedRunOrigins(input).root);
});

for (const connected of [false, true])
  test(`part-wide live revisions protect resolved proofs across paragraphs; related scope=${connected}`, () => {
    const liveScope = connected ? other : scope;
    const liveNs = connected ? otherNs : ns;
    const related = connected
      ? `<w:r xmlns:q="${otherNs}" q:origin="${other}/${source}" q:identity="source"><w:t>X</w:t></w:r>`
      : '';
    const input = part(
      completed +
        `<w:p>${related}<w:ins xmlns:q="${liveNs}" q:scope="${liveScope}" w:id="1" w:author="Ada"><w:r><w:t>Live</w:t></w:r></w:ins></w:p>`
    );
    expect(projectCompletedRunOrigins(input)).toBe(input);
    expect(serializeOoxmlPart(input)).toContain('resolved="1"');
  });

test('inactive proofs without explicit resolution and property history snapshots remain literal', () => {
  const inactive = part(`<w:p>${fragment('A', false)}${fragment('B', false)}</w:p>`);
  expect(projectCompletedRunOrigins(inactive)).toBe(inactive);
  const snapshot = part(
    `<w:p><w:pPr><w:pPrChange w:id="2" w:author="Ada"><w:pPr><w:opaque>${completed}</w:opaque></w:pPr></w:pPrChange></w:pPr></w:p>${completed}`
  );
  const output = serializeOoxmlPart(snapshot);
  expect(output).toContain('resolved="1"');
  expect((output.match(/resolved="1"/g) ?? []).length).toBe(4);
  expect(canonicalOoxmlFingerprint(projectCompletedRunOrigins(snapshot))).not.toBe(
    canonicalOoxmlFingerprint(snapshot)
  );
});

test('a retained alias blocks its whole scope even when another source uses another alias', () => {
  const first = fragment('A');
  const second = fragment('B')
    .replaceAll('xmlns:o=', 'xmlns:alias=')
    .replaceAll(' o:', ' alias:')
    .replaceAll(source, '22345678-1234-8234-a234-123456789abc');
  const input = part(
    `<w:p>${first}${second}</w:p><w:p><w:ins xmlns:q="${otherNs}" q:scope="${other}" q:retained='["o"]' w:id="1" w:author="Ada"><w:r><w:t>Live</w:t></w:r></w:ins></w:p>`
  );
  expect(projectCompletedRunOrigins(input)).toBe(input);
});
