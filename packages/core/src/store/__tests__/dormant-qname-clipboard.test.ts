import { expect, test } from 'bun:test';
import { part, paragraphId, apply } from './tracked-edit-fixture.ts';
import { findNode, readOoxmlPart, serializeOoxmlPart } from '../index.ts';
import { namespaceBindingsAt } from '../package/wml-namespace.ts';
import { RUN_ORIGIN_NS, originNodes } from '../store/run-origin-metadata.ts';
const scope = '11111111-1111-4111-8111-111111111111';
for (const placement of ['run', 'paragraph', 'revision'] as const)
  for (const prefix of ['o', 'č'])
    for (const active of [false, true])
      test(`copied provenance preserves foreign QName alias ${prefix}/${active}/${placement}`, () => {
        const body = `<w:r ${placement === 'run' ? `xmlns:${prefix}="${RUN_ORIGIN_NS}"` : ''} xmlns:keep="urn:keep" ${prefix}:origin="${scope}/22222222-2222-4222-8222-222222222222" keep:type="${prefix}:Value"><w:rPr/><w:t>A</w:t></w:r>`;
        const source = part(
          `<w:p ${placement !== 'run' && (!active || placement === 'paragraph') ? `xmlns:${prefix}="${RUN_ORIGIN_NS}"` : ''}>${active ? `<w:ins ${placement !== 'paragraph' ? `xmlns:${prefix}="${RUN_ORIGIN_NS}"` : ''} w:id="9" w:author="Other" ${prefix}:scope="${scope}">${body}</w:ins>` : body}</w:p>`
        );
        const paragraph = findNode(source, paragraphId(source))!;
        const host = part('<w:p><w:r><w:t>B</w:t></w:r></w:p>');
        let pasted = apply(host, {
          op: 'insertFragment',
          paragraphId: paragraphId(host),
          offset: 1,
          blocks: [paragraph],
        });
        if (active) pasted = apply(pasted, { op: 'acceptAllRevisions' });
        for (const current of [
          pasted,
          (() => {
            const r = readOoxmlPart(serializeOoxmlPart(pasted), {
              name: pasted.name,
              contentType: pasted.contentType,
            });
            if (!r.ok) throw Error(r.reason);
            return r.part;
          })(),
        ]) {
          const run = originNodes(current.root).find(
            (n) =>
              n.kind === 'run' &&
              n.attributes.some((a) => a.namespaceUri === 'urn:keep' && a.localName === 'type')
          )!;
          if (run.kind !== 'run') throw Error('Missing copied run.');
          expect(namespaceBindingsAt(current, run).get(prefix)).toBe(RUN_ORIGIN_NS);
        }
      });

test('ordinary copies retain default namespace used by opaque QName data', () => {
  const source = part(
    `<w:p><w:r xmlns="${RUN_ORIGIN_NS}" xmlns:o="${RUN_ORIGIN_NS}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" o:origin="${scope}/22222222-2222-4222-8222-222222222222" xsi:type="OpaqueType"><w:t>A</w:t></w:r></w:p>`
  );
  const p = findNode(source, paragraphId(source))!,
    host = part('<w:p><w:r><w:t>B</w:t></w:r></w:p>');
  const pasted = apply(host, {
    op: 'insertFragment',
    paragraphId: paragraphId(host),
    offset: 1,
    blocks: [p],
  });
  const r = readOoxmlPart(serializeOoxmlPart(pasted), {
    name: pasted.name,
    contentType: pasted.contentType,
  });
  if (!r.ok) throw Error(r.reason);
  for (const current of [pasted, r.part]) {
    const run = originNodes(current.root).find(
      (n) =>
        n.kind === 'run' &&
        n.attributes.some((a) => a.namespaceUri === 'http://www.w3.org/2001/XMLSchema-instance')
    )!;
    if (run.kind !== 'run') throw Error('Missing copied run.');
    expect(namespaceBindingsAt(current, run).get('')).toBe(RUN_ORIGIN_NS);
  }
});
