import { expect, test } from 'bun:test';
import {
  readOoxmlPart,
  canonicalOoxmlFingerprint,
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  type OoxmlPart,
} from '../index.ts';
import { paragraphId, apply, ADA } from './tracked-edit-fixture.ts';
import { RUN_ORIGIN_NS, originNamespaceForScope } from '../store/run-origin-metadata.ts';
const scope = '11111111-1111-4111-8111-111111111111',
  other = '33333333-3333-4333-8333-333333333333';
const token = `${scope}/22222222-2222-4222-8222-222222222222`;
const fp = canonicalOoxmlFingerprint;
function load(content: string, uri: string): OoxmlPart {
  const r = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="${uri}"><w:body>${content}</w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!r.ok) throw Error(r.reason);
  return r.part;
}
for (const [uri, attributes] of [
  [originNamespaceForScope(scope), 'o:origin="invalid"'],
  [originNamespaceForScope(scope), `o:origin="${token}" o:future="opaque"`],
  [originNamespaceForScope(other), `o:origin="${token}"`],
  [originNamespaceForScope(scope), `o:origin="${token}" o:absent="false"`],
  [originNamespaceForScope(scope), `o:origin="${token}" o:retained="[0]"`],
] as const)
  test(`invalid versioned provenance stays opaque: ${uri}/${attributes}`, () => {
    const content = `<w:r ${attributes}><w:rPr/><w:t>A</w:t></w:r>`;
    const source = load(
      `<w:p><w:ins w:id="7" w:author="Other" o:scope="${scope}">${content}</w:ins></w:p>`,
      uri
    );
    const decided = apply(source, { op: 'acceptAllRevisions' }),
      golden = load(`<w:p>${content}</w:p>`, uri);
    expect(fp(decided)).toBe(fp(golden));
    expect(() =>
      prepareTrackedRunOrigins(golden, paragraphId(golden), [{ start: 0, end: 1 }])
    ).toThrow('opaque');
  });
for (const uri of [`${RUN_ORIGIN_NS}/future`, `${RUN_ORIGIN_NS}:future`])
  test(`future namespace data survives a new complete proposal: ${uri}`, () => {
    const source = load(`<w:p><w:r o:origin="${token}"><w:rPr/><w:t>AB</w:t></w:r></w:p>`, uri),
      id = paragraphId(source),
      before = fp(source);
    const prepared = prepareTrackedRunOrigins(source, id, [{ start: 1, end: 1 }]);
    const pending = sealTrackedRunOrigins(
      apply(prepared.part, {
        op: 'insertText',
        paragraphId: id,
        offset: 1,
        text: 'X',
        revision: ADA,
      }),
      id,
      prepared.scope!,
      source
    );
    const expected = apply(source, { op: 'insertText', paragraphId: id, offset: 1, text: 'X' });
    expect(fp(apply(pending, { op: 'acceptAllRevisions' }))).toBe(fp(expected));
    expect(fp(apply(pending, { op: 'rejectAllRevisions' }))).toBe(before);
    expect(fp(source)).toBe(before);
  });
