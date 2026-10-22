import { expect, test } from 'bun:test';
import { part, apply } from './tracked-edit-fixture.ts';
import { readOoxmlPart, serializeOoxmlPart, type OoxmlPart, type OoxmlNode } from '../index.ts';
import { namespaceBindingsAt } from '../package/wml-namespace.ts';
const walk = (p: OoxmlPart) => {
  const nodes: OoxmlNode[] = [p.root];
  const all: OoxmlNode[] = [];
  while (nodes.length) {
    const n = nodes.pop()!;
    all.push(n);
    if (n.kind !== 'textValue') for (const c of n.children) nodes.push(c);
  }
  return all;
};
for (const prefix of ['o', 'č'])
  for (const kind of ['ins', 'del'])
    for (const saved of [false, true])
      test(`revision removal preserves inherited QName aliases ${prefix}/${kind}/${saved}`, () => {
        const text = kind === 'ins' ? 't' : 'delText';
        const source = part(
          `<w:p xmlns:${prefix}="urn:outer"><w:${kind} xmlns:${prefix}="urn:inner" xmlns:unused="urn:unused" w:id="1" w:author="Other"><w:r xmlns:keep="urn:keep" keep:name="inherited" keep:type="${prefix}:Type"><w:${text}>A</w:${text}></w:r><w:r xmlns:${prefix}="urn:shadow" xmlns:keep="urn:keep" keep:name="shadow" keep:type="${prefix}:Type"><w:${text}>B</w:${text}></w:r></w:${kind}></w:p>`
        );
        let current = apply(source, {
          op: kind === 'ins' ? 'acceptAllRevisions' : 'rejectAllRevisions',
        });
        if (saved) {
          const r = readOoxmlPart(serializeOoxmlPart(current), {
            name: current.name,
            contentType: current.contentType,
          });
          if (!r.ok) throw Error(r.reason);
          current = r.part;
        }
        for (const name of ['inherited', 'shadow']) {
          const node = walk(current).find(
            (n) =>
              n.kind === 'run' &&
              n.attributes.some(
                (a) => a.namespaceUri === 'urn:keep' && a.localName === 'name' && a.value === name
              )
          )!;
          if (node.kind !== 'run') throw Error('Missing surviving run.');
          expect(namespaceBindingsAt(current, node).get(prefix)).toBe(
            name === 'inherited' ? 'urn:inner' : 'urn:shadow'
          );
        }
        expect(serializeOoxmlPart(current)).not.toContain('urn:unused');
      });
for (const kind of ['ins', 'del'])
  for (const saved of [false, true])
    test(`revision removal preserves default QName context ${kind}/${saved}`, () => {
      const text = kind === 'ins' ? 't' : 'delText';
      const source = part(
        `<w:p xmlns="urn:outer"><w:${kind} xmlns="urn:inner" w:id="1" w:author="Other"><w:r xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="OpaqueType"><w:${text}>A</w:${text}></w:r></w:${kind}></w:p>`
      );
      let current = apply(source, {
        op: kind === 'ins' ? 'acceptAllRevisions' : 'rejectAllRevisions',
      });
      if (saved) {
        const r = readOoxmlPart(serializeOoxmlPart(current), {
          name: current.name,
          contentType: current.contentType,
        });
        if (!r.ok) throw Error(r.reason);
        current = r.part;
      }
      const run = walk(current).find((n) => n.kind === 'run')!;
      expect(namespaceBindingsAt(current, run).get('')).toBe('urn:inner');
    });
