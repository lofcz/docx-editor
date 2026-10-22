import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  prepareTrackedRunOrigins,
  readOoxmlPart,
} from '../index.ts';
import { RUN_ORIGIN_NS } from '../store/run-origin-metadata.ts';
import { apply, paragraphId } from './tracked-edit-fixture.ts';
import { applyTreeOp } from '../store/tree-ops.ts';
const scope = '11111111-1111-4111-8111-111111111111';
const origin = `${scope}/22222222-2222-4222-8222-222222222222`;
const run = (attrs: string) => `<w:r ${attrs}><w:rPr/><w:t>A</w:t></w:r>`;
function load(body: string) {
  const read = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="${RUN_ORIGIN_NS}" xmlns:future="${RUN_ORIGIN_NS}:future"><w:body>${body}</w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!read.ok) throw Error(read.reason);
  return read.part;
}
for (const attributes of [
  'o:origin="invalid"',
  `o:origin="${origin}" o:absent="false"`,
  `o:origin="${origin}" o:future="untouched"`,
  `o:origin="${origin}" o:retained="{}"`,
  `o:origin="${origin}" o:retained="[0]"`,
  `o:origin="${origin}" o:retained="[&quot;a&quot;,&quot;a&quot;]"`,
  `future:origin="${origin}"`,
])
  test(`malformed and future run-origin metadata stays opaque: ${attributes}`, () => {
    const content = run(attributes) + run(attributes);
    const source = load(
      `<w:p><w:ins w:id="7" w:author="Other" o:scope="${scope}">${content}</w:ins></w:p>`
    );
    const golden = load(`<w:p>${content}</w:p>`);
    const decided = apply(source, { op: 'acceptAllRevisions' });
    expect(canonicalOoxmlFingerprint(decided)).toBe(canonicalOoxmlFingerprint(golden));
    if (!attributes.startsWith('future:'))
      expect(() =>
        prepareTrackedRunOrigins(golden, paragraphId(golden), [{ start: 0, end: 1 }])
      ).toThrow('already carries');
  });

for (const extra of ['o:scope="invalid"', `o:scope="${scope}" o:future="untouched"`])
  test(`unrecognized native scope metadata cannot normalize source runs: ${extra}`, () => {
    const content = run(`o:origin="${origin}"`) + run(`o:origin="${origin}"`);
    const source = load(`<w:p><w:ins w:id="7" w:author="Other" ${extra}>${content}</w:ins></w:p>`);
    const decided = apply(source, { op: 'acceptAllRevisions' });
    const expected = load(`<w:p>${content}</w:p>`);
    expect(canonicalOoxmlFingerprint(decided)).toBe(canonicalOoxmlFingerprint(expected));
  });

test('a scoped paragraph-mark revision cannot claim ordinary run provenance', () => {
  const content = run(`o:origin="${origin}"`) + run(`o:origin="${origin}"`);
  const source = load(
    `<w:p><w:pPr><w:rPr><w:ins w:id="7" w:author="Other" o:scope="${scope}"/></w:rPr></w:pPr>${content}</w:p>`
  );
  const decided = apply(source, { op: 'acceptAllRevisions' });
  const p = findNode(decided, paragraphId(decided))!;
  if (p.kind !== 'paragraph') throw Error('paragraph');
  expect(p.children.filter((node) => node.kind === 'run')).toHaveLength(2);
  for (const node of p.children.filter((node) => node.kind === 'run'))
    expect(node.attributes.some((attr) => attr.namespaceUri === RUN_ORIGIN_NS)).toBe(true);
});

for (const textAttributes of [
  'o:textOrigin="invalid" o:space="absent"',
  `o:textOrigin="${origin}" o:space="future"`,
  `o:textOrigin="${origin}" o:space="absent" o:future="opaque"`,
])
  test(`unknown text-origin metadata stays opaque: ${textAttributes}`, () => {
    const content = `<w:r o:origin="${origin}"><w:rPr/><w:t xml:space="preserve" ${textAttributes}>A</w:t></w:r>`;
    const source = load(
      `<w:p><w:ins w:id="7" w:author="Other" o:scope="${scope}">${content}</w:ins></w:p>`
    );
    const expected = load(`<w:p>${content.replace(` o:origin="${origin}"`, '')}</w:p>`);
    expect(canonicalOoxmlFingerprint(apply(source, { op: 'acceptAllRevisions' }))).toBe(
      canonicalOoxmlFingerprint(expected)
    );
  });

test('text-origin metadata cannot claim a run without matching provenance', () => {
  const content = `<w:r><w:t o:textOrigin="${origin}" o:space="absent" xml:space="preserve">A</w:t></w:r>`;
  const source = load(
    `<w:p><w:ins w:id="7" w:author="Other" o:scope="${scope}">${content}</w:ins></w:p>`
  );
  expect(canonicalOoxmlFingerprint(apply(source, { op: 'acceptAllRevisions' }))).toBe(
    canonicalOoxmlFingerprint(load(`<w:p>${content}</w:p>`))
  );
});

test('oversized imported provenance returns a native refusal without changing the source', () => {
  const source = load(
    `<w:p><w:ins w:id="7" w:author="Other" o:scope="${scope}">${run(`o:origin="${origin}"`)}</w:ins>${'<w:proofErr w:type="spellStart"/>'.repeat(100_001)}</w:p>`
  );
  const before = canonicalOoxmlFingerprint(source);
  expect(applyTreeOp(source, { op: 'acceptAllRevisions' })).toMatchObject({
    ok: false,
    reason: 'resource-limit',
  });
  expect(canonicalOoxmlFingerprint(source)).toBe(before);
}, 15000);
