import { test, expect } from 'bun:test';
import { part, paragraphId, apply, ADA } from './tracked-edit-fixture.ts';
import {
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  collectRevisionSites,
  serializeOoxmlPart,
  readOoxmlPart,
  canonicalOoxmlFingerprint,
  WML_NAMESPACE_URI,
  type OoxmlPart,
} from '../index.ts';
import { readRevisionOrigin, RUN_ORIGIN_NS } from '../store/run-origin-metadata.ts';
const fp = canonicalOoxmlFingerprint;
function prepared() {
  const source = part(
      '<w:p><w:r><w:rPr/><w:t>ABC</w:t></w:r><w:r><w:rPr/><w:t>DEF</w:t></w:r></w:p>'
    ),
    id = paragraphId(source);
  const a = prepareTrackedRunOrigins(source, id, [{ start: 1, end: 1 }]);
  let current = sealTrackedRunOrigins(
    apply(a.part, { op: 'insertText', paragraphId: id, offset: 1, text: 'X', revision: ADA }),
    id,
    a.scope!,
    source
  );
  const before = current,
    b = prepareTrackedRunOrigins(current, id, [{ start: 6, end: 6 }]);
  current = sealTrackedRunOrigins(
    apply(b.part, {
      op: 'insertText',
      paragraphId: id,
      offset: 6,
      text: 'Y',
      revision: { author: 'Bob', date: '2026-10-22T12:00:00Z' },
    }),
    id,
    b.scope!,
    before
  );
  return { source, id, current, scopes: [a.scope!, b.scope!] };
}
function reopen(p: OoxmlPart) {
  const r = readOoxmlPart(serializeOoxmlPart(p), { name: p.name, contentType: p.contentType });
  if (!r.ok) throw Error(r.reason);
  return r.part;
}
function resolve(p: OoxmlPart, scope: string, action: 'accept' | 'reject') {
  const site = collectRevisionSites(p).find((s) => readRevisionOrigin(s.node)?.scope === scope)!;
  const value = (name: string) =>
    site.node.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)
      ?.value;
  return apply(p, {
    op: action === 'accept' ? 'acceptRevision' : 'rejectRevision',
    revision: { id: value('id')!, author: value('author')!, date: value('date') },
    siteNodeIds: [site.node.id],
  });
}
function golden(source: OoxmlPart, a: 'accept' | 'reject', b: 'accept' | 'reject') {
  const id = paragraphId(source);
  let current = source;
  if (a === 'accept')
    current = apply(current, { op: 'insertText', paragraphId: id, offset: 1, text: 'X' });
  if (b === 'accept')
    current = apply(current, {
      op: 'insertText',
      paragraphId: id,
      offset: a === 'accept' ? 6 : 5,
      text: 'Y',
    });
  return current;
}
for (const saved of [false, true])
  for (const order of [
    [0, 1],
    [1, 0],
  ] as const)
    for (const a of ['accept', 'reject'] as const)
      for (const b of ['accept', 'reject'] as const)
        test(`multiple scopes resolve independently ${saved}/${order}/${a}/${b}`, () => {
          const f = prepared();
          let current = saved ? reopen(f.current) : f.current;
          const other = collectRevisionSites(current)
            .filter((s) => readRevisionOrigin(s.node)?.scope === f.scopes[order[1]])
            .map((s) => fp(s.node));
          current = resolve(current, f.scopes[order[0]]!, order[0] === 0 ? a : b);
          expect(
            collectRevisionSites(current)
              .filter((s) => readRevisionOrigin(s.node)?.scope === f.scopes[order[1]])
              .map((s) => fp(s.node))
          ).toEqual(other);
          current = resolve(current, f.scopes[order[1]]!, order[1] === 0 ? a : b);
          expect(fp(current)).toBe(fp(golden(f.source, a, b)));
          expect(serializeOoxmlPart(current)).not.toContain(RUN_ORIGIN_NS);
        });
for (const saved of [false, true])
  for (const action of ['accept', 'reject'] as const)
    test(`multiple scopes settle in one decision ${saved}/${action}`, () => {
      const f = prepared();
      let current = saved ? reopen(f.current) : f.current;
      current = apply(current, {
        op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
      });
      expect(fp(current)).toBe(fp(golden(f.source, action, action)));
      expect(serializeOoxmlPart(current)).not.toContain(RUN_ORIGIN_NS);
    });
