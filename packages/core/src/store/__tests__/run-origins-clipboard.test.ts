import { expect, test } from 'bun:test';
import { part, paragraphId, apply, ADA } from './tracked-edit-fixture.ts';
import {
  findNode,
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  type OoxmlPart,
  type OoxmlParagraphNode,
} from '../index.ts';
import { applyTreeOp } from '../store/tree-ops.ts';
import { freshenCopiedRunOrigins } from '../store/run-origin-clipboard.ts';
import {
  readRevisionOrigin,
  readRunOrigin,
  originNodes,
  RUN_ORIGIN_NS,
} from '../store/run-origin-metadata.ts';
const fingerprint = canonicalOoxmlFingerprint;
function paragraph(source: OoxmlPart): OoxmlParagraphNode {
  const node = findNode(source, paragraphId(source));
  if (node?.kind !== 'paragraph') throw Error('Missing paragraph.');
  return node;
}
function reopen(source: OoxmlPart) {
  const read = readOoxmlPart(serializeOoxmlPart(source), {
    name: source.name,
    contentType: source.contentType,
  });
  if (!read.ok) throw Error(read.reason);
  return read.part;
}
function fixture(properties: string) {
  const source = part(
    `<w:p><w:r xmlns:keep="urn:keep" keep:run="original">${properties}<w:t keep:text="original">AB</w:t></w:r></w:p>`
  );
  const id = paragraphId(source);
  const prepared = prepareTrackedRunOrigins(source, id, [{ start: 1, end: 1 }]);
  const inserted = apply(prepared.part, {
    op: 'insertText',
    paragraphId: id,
    offset: 1,
    text: 'X',
    revision: ADA,
  });
  const pending = sealTrackedRunOrigins(inserted, id, prepared.scope!, source);
  return { source, id, pending, scope: prepared.scope! };
}
for (const properties of ['', '<w:rPr/>'])
  for (const saved of [false, true])
    for (const action of ['accept', 'reject'] as const)
      for (const copies of [1, 2])
        test(`copied original runs remain independent: ${properties}/${saved}/${action}/${copies}`, () => {
          const { source, id, pending } = fixture(properties);
          const sourceBefore = fingerprint(source),
            pendingBefore = fingerprint(pending);
          const original = paragraph(pending).children.find((node) => node.kind === 'run');
          if (!original) throw Error('Missing original run.');
          const template = paragraph(
            part(
              `<w:p><w:r xmlns:keep="urn:keep" keep:run="original">${properties}<w:t keep:text="original">A</w:t></w:r></w:p>`
            )
          );
          let pasted = pending;
          for (let index = 0; index < copies; index++)
            pasted = apply(pasted, {
              op: 'insertFragment',
              paragraphId: id,
              offset: 3 + index,
              blocks: [{ ...template, children: [original] }],
            });
          if (saved) pasted = reopen(pasted);
          const decided = apply(pasted, {
            op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
          });
          let golden =
            action === 'accept'
              ? apply(source, { op: 'insertText', paragraphId: id, offset: 1, text: 'X' })
              : source;
          for (let index = 0; index < copies; index++)
            golden = apply(golden, {
              op: 'insertFragment',
              paragraphId: id,
              offset: (action === 'accept' ? 3 : 2) + index,
              blocks: [template],
            });
          expect(fingerprint(decided)).toBe(fingerprint(golden));
          expect(fingerprint(source)).toBe(sourceBefore);
          expect(fingerprint(pending)).toBe(pendingBefore);
        });

test('complete copies receive independent scope and source tokens', () => {
  const { pending, scope } = fixture('<w:rPr/>');
  const original = paragraph(pending),
    before = fingerprint(original);
  const first = freshenCopiedRunOrigins([original], 500_000, 64);
  const second = freshenCopiedRunOrigins([original], 500_000, 64);
  const one = originNodes(first[0]!)
    .map((node) => readRevisionOrigin(node))
    .filter((site) => site !== null);
  const two = originNodes(second[0]!)
    .map((node) => readRevisionOrigin(node))
    .filter((site) => site !== null);
  expect(one.length).toBeGreaterThan(0);
  expect(one.every((site) => site.scope !== scope)).toBe(true);
  expect(two[0]!.scope).not.toBe(one[0]!.scope);
  for (const node of originNodes(first[0]!)) {
    const origin = readRunOrigin(node);
    if (origin) expect(origin.scope).toBe(one[0]!.scope);
  }
  expect(fingerprint(original)).toBe(before);
});

test('ordinary clipboard fragments preserve object identity', () => {
  const blocks = [paragraph(part('<w:p><w:r><w:t>A</w:t></w:r></w:p>'))];
  expect(freshenCopiedRunOrigins(blocks, 500_000, 64)).toBe(blocks);
});

test('future clipboard provenance remains opaque', () => {
  const source = paragraph(
    part(
      `<w:p xmlns:o="${RUN_ORIGIN_NS}"><w:r o:origin="future"><w:t o:textOrigin="future">A</w:t></w:r></w:p>`
    )
  );
  const before = fingerprint(source);
  expect(fingerprint(freshenCopiedRunOrigins([source], 500_000, 64)[0]!)).toBe(before);
  expect(fingerprint(source)).toBe(before);
});

test('clipboard text provenance requires matching run ownership', () => {
  const scope = '11111111-1111-4111-8111-111111111111';
  const source = paragraph(
    part(
      `<w:p xmlns:o="${RUN_ORIGIN_NS}"><w:r><w:t o:textOrigin="${scope}/22222222-2222-4222-8222-222222222222" o:space="absent" xml:space="preserve">A</w:t></w:r></w:p>`
    )
  );
  const before = fingerprint(source);
  expect(fingerprint(freshenCopiedRunOrigins([source], 500_000, 64)[0]!)).toBe(before);
  expect(fingerprint(source)).toBe(before);
});

test('oversized copied provenance returns a native refusal without changing the host', () => {
  const scope = '11111111-1111-4111-8111-111111111111';
  const fragment = paragraph(
    part(
      `<w:p xmlns:o="${RUN_ORIGIN_NS}"><w:r o:origin="${scope}/22222222-2222-4222-8222-222222222222">${'<w:t>A</w:t>'.repeat(50_001)}</w:r></w:p>`
    )
  );
  const host = part('<w:p><w:r><w:t>B</w:t></w:r></w:p>'),
    before = fingerprint(host);
  expect(
    applyTreeOp(host, {
      op: 'insertFragment',
      paragraphId: paragraphId(host),
      offset: 1,
      blocks: [fragment],
    })
  ).toMatchObject({ ok: false, reason: 'resource-limit' });
  expect(fingerprint(host)).toBe(before);
}, 15000);

for (const action of ['accept', 'reject'] as const)
  for (const saved of [false, true])
    test(`copy retains completed peer proof for a pending concurrent proposal: ${action}/${saved}`, () => {
      const a = '11111111-1111-4111-8111-111111111111',
        b = '22222222-2222-4222-8222-222222222222';
      const runId = '33333333-3333-4333-8333-333333333333',
        textId = '44444444-4444-4444-8444-444444444444';
      const attrs = `xmlns:a="${RUN_ORIGIN_NS}/${a}" xmlns:b="${RUN_ORIGIN_NS}/${b}"`;
      const marked = (prefix: 'a' | 'b', text: string) => {
        const scope = prefix === 'a' ? a : b,
          resolved = prefix === 'b' ? ` ${prefix}:resolved="1"` : '';
        return `<w:r ${prefix}:origin="${scope}/${runId}" ${prefix}:identity="source"${resolved}><w:rPr/><w:t xml:space="preserve" ${prefix}:textOrigin="${scope}/${textId}" ${prefix}:identity="source" ${prefix}:space="preserve"${resolved}>${text}</w:t></w:r>`;
      };
      const plain = (text: string) => `<w:r><w:rPr/><w:t xml:space="preserve">${text}</w:t></w:r>`;
      const pending = part(
        `<w:p ${attrs}>${marked('b', 'Alpha end')}<w:ins w:id="1" w:author="Alice" a:scope="${a}">${marked('a', ' ')}</w:ins><w:hyperlink w:anchor="one"><w:ins w:id="2" w:author="Alice" a:scope="${a}">${marked('a', 'one')}</w:ins></w:hyperlink>${marked('b', ' ')}<w:hyperlink w:anchor="two">${marked('b', 'two')}</w:hyperlink></w:p>`
      );
      const copied = freshenCopiedRunOrigins([paragraph(pending)], 500_000, 64)[0]!;
      if (copied.kind !== 'paragraph') throw Error('Missing copied paragraph');
      const original = paragraph(pending);
      const copiedPart = {
        ...pending,
        root: {
          ...pending.root,
          children: pending.root.children.map((node) =>
            node.kind === 'body'
              ? {
                  ...node,
                  children: node.children.map((child) =>
                    child.id === original.id ? copied : child
                  ),
                }
              : node
          ),
        },
      } as OoxmlPart;
      const input = saved ? reopen(copiedPart) : copiedPart;
      const result = apply(input, {
        op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
      });
      const golden = part(
        `<w:p ${attrs}>${plain('Alpha end ')}${action === 'accept' ? `<w:hyperlink w:anchor="one">${plain('one')}</w:hyperlink>${plain(' ')}` : ''}<w:hyperlink w:anchor="two">${plain('two')}</w:hyperlink></w:p>`
      );
      expect(fingerprint(result)).toBe(fingerprint(golden));
    });
