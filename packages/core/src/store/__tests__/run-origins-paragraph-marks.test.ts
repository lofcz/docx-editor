import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  bodyStoryRoot,
  collectRevisionSites,
  findNode,
  prepareTrackedRunOrigins,
  readOoxmlPart,
  sealTrackedRunOrigins,
  serializeOoxmlPart,
  storyParagraphs,
  WML_NAMESPACE_URI,
} from '../index';
import { ADA, apply, paragraphId, part } from './tracked-edit-fixture';
import { freshenCopiedRunOrigins } from '../store/run-origin-clipboard';
import { originNodes, readRevisionOrigins, readRunOrigins } from '../store/run-origin-metadata';
import { runWithTransactionActor } from '../package/actor-scoped-ids';
import { retiredParagraphExportPackage, isJoinedParagraph } from '../package/retired-paragraph';
import { singlePartPackage } from '../package/ooxml-package';

for (const offsets of [[2], [0], [4], [0, 2, 2, 4]])
  test(`paragraph-mark provenance restores exact source XML after rejection: ${offsets.join('/')}`, () => {
    const source = part(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">ABCD</w:t></w:r></w:p>'
    );
    const id = paragraphId(source);
    const original = serializeOoxmlPart(source);
    const prepared = prepareTrackedRunOrigins(
      source,
      id,
      offsets.map((offset) => ({ start: offset, end: offset }))
    );
    expect(prepared.scope).not.toBeNull();
    let pending = apply(prepared.part, {
      op: 'splitParagraphMany',
      paragraphId: id,
      offsets,
      revision: ADA,
    });
    pending = sealTrackedRunOrigins(pending, id, prepared.scope!, source);
    for (const paragraph of storyParagraphs(bodyStoryRoot(pending)!)) {
      if (paragraph.id !== id)
        pending = sealTrackedRunOrigins(pending, paragraph.id, prepared.scope!, source);
    }
    const parsed = readOoxmlPart(serializeOoxmlPart(pending), {
      name: pending.name,
      contentType: pending.contentType,
    });
    if (!parsed.ok) throw Error(parsed.reason);
    for (const document of [pending, parsed.part]) {
      const rejected = apply(document, { op: 'rejectAllRevisions' });
      expect(serializeOoxmlPart(rejected)).toBe(original);
      expect(canonicalOoxmlFingerprint(rejected)).toBe(canonicalOoxmlFingerprint(source));
      const accepted = apply(document, { op: 'acceptAllRevisions' });
      const direct = apply(source, { op: 'splitParagraphMany', paragraphId: id, offsets });
      expect(canonicalOoxmlFingerprint(accepted)).toBe(canonicalOoxmlFingerprint(direct));
    }
  });

for (const offsets of [[2], [0], [4], [0, 2, 2, 4]])
  test(`shared paragraph rejection retains live source slots and exports exact original XML: ${offsets}`, () => {
    const source = part(
      '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="12345678"><w:r><w:rPr><w:b/></w:rPr><w:t>ABCD</w:t></w:r></w:p>'
    );
    const id = paragraphId(source);
    const prepared = prepareTrackedRunOrigins(
      source,
      id,
      offsets.map((offset) => ({ start: offset, end: offset }))
    );
    let pending = apply(prepared.part, {
      op: 'splitParagraphMany',
      paragraphId: id,
      offsets,
      revision: ADA,
    });
    for (const paragraph of storyParagraphs(bodyStoryRoot(pending)!))
      pending = sealTrackedRunOrigins(pending, paragraph.id, prepared.scope!, source);
    const ids = storyParagraphs(bodyStoryRoot(pending)!).map((paragraph) => paragraph.id);
    const rejected = runWithTransactionActor('alice', () =>
      apply(pending, { op: 'rejectAllRevisions' })
    );
    const physical = bodyStoryRoot(rejected)!.children.filter((node) => node.kind === 'paragraph');
    expect(physical.map((paragraph) => paragraph.id)).toEqual(ids);
    expect(physical.filter(isJoinedParagraph)).toHaveLength(offsets.length);
    const raw = serializeOoxmlPart(rejected);
    expect(raw).toContain('retired-paragraph:2');
    if (offsets.some((offset) => offset > 0 && offset < 4))
      expect(raw).toContain('tracked-run-origins:1/');
    const read = readOoxmlPart(raw, { name: source.name, contentType: source.contentType });
    if (!read.ok) throw Error(read.reason);
    for (const live of [rejected, read.part]) {
      const before = canonicalOoxmlFingerprint(live);
      const exported = retiredParagraphExportPackage(singlePartPackage(live)).parts.get(live.name)!;
      expect(serializeOoxmlPart(exported)).toBe(serializeOoxmlPart(source));
      expect(canonicalOoxmlFingerprint(live)).toBe(before);
    }
  });

test('a split fragment cannot seal its source proof before the fresh paragraph mark witnesses that scope', () => {
  const source = part('<w:p><w:r><w:t>ABCD</w:t></w:r></w:p>');
  const id = paragraphId(source);
  const prepared = prepareTrackedRunOrigins(source, id, [{ start: 2, end: 2 }]);
  const split = apply(prepared.part, {
    op: 'splitParagraphMany',
    paragraphId: id,
    offsets: [2],
    revision: ADA,
  });
  const tail = storyParagraphs(bodyStoryRoot(split)!).find((node) => node.id !== id)!;
  expect(() => sealTrackedRunOrigins(split, tail.id, prepared.scope!, source)).toThrow(
    'new native revision sites'
  );
  const witnessed = sealTrackedRunOrigins(split, id, prepared.scope!, source);
  const sealed = sealTrackedRunOrigins(witnessed, tail.id, prepared.scope!, source);
  expect(findNode(sealed, tail.id)).not.toBeNull();
  expect(serializeOoxmlPart(apply(sealed, { op: 'rejectAllRevisions' }))).toBe(
    serializeOoxmlPart(source)
  );
});

for (const action of ['accept', 'reject'] as const)
  test(`copied paragraph marks own independent source proofs through ${action} and reopen`, () => {
    const source = part('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ABCD</w:t></w:r></w:p>');
    const id = paragraphId(source);
    const prepared = prepareTrackedRunOrigins(source, id, [{ start: 2, end: 2 }]);
    let pending = apply(prepared.part, {
      op: 'splitParagraphMany',
      paragraphId: id,
      offsets: [2],
      revision: ADA,
    });
    for (const paragraph of storyParagraphs(bodyStoryRoot(pending)!))
      pending = sealTrackedRunOrigins(pending, paragraph.id, prepared.scope!, source);
    const before = canonicalOoxmlFingerprint(pending);
    const paragraphs = storyParagraphs(bodyStoryRoot(pending)!);
    const copied = freshenCopiedRunOrigins(paragraphs, 500_000, 64);
    const copies = copied.flatMap(originNodes);
    const scopes = copies.flatMap(readRevisionOrigins).map((origin) => origin.scope);
    expect(scopes).toHaveLength(1);
    expect(scopes[0]).not.toBe(prepared.scope);
    expect(copies.flatMap(readRunOrigins).every((origin) => origin.scope === scopes[0])).toBe(true);
    const copiedPart = {
      ...source,
      root: {
        ...source.root,
        children: source.root.children.map((node) =>
          node.kind === 'body' ? { ...node, children: copied } : node
        ),
      },
    };
    const reopened = readOoxmlPart(serializeOoxmlPart(copiedPart), {
      name: pending.name,
      contentType: pending.contentType,
    });
    if (!reopened.ok) throw Error(reopened.reason);
    const decided = apply(reopened.part, {
      op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
    });
    const expected =
      action === 'accept'
        ? apply(source, { op: 'splitParagraphMany', paragraphId: id, offsets: [2] })
        : source;
    expect(serializeOoxmlPart(decided)).toBe(serializeOoxmlPart(expected));
    expect(canonicalOoxmlFingerprint(pending)).toBe(before);
  });

for (const order of [
  [0, 1],
  [1, 0],
])
  for (const decisions of [
    ['accept', 'accept'],
    ['reject', 'reject'],
    ['accept', 'reject'],
    ['reject', 'accept'],
  ] as const)
    test(`paragraph source proofs survive separate decisions and reopen: ${order}/${decisions}`, () => {
      const source = part('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ABCD</w:t></w:r></w:p>');
      const id = paragraphId(source);
      const prepared = prepareTrackedRunOrigins(source, id, [
        { start: 1, end: 1 },
        { start: 3, end: 3 },
      ]);
      let pending = apply(prepared.part, {
        op: 'splitParagraphMany',
        paragraphId: id,
        offsets: [1, 3],
        revision: ADA,
      });
      for (const paragraph of storyParagraphs(bodyStoryRoot(pending)!))
        pending = sealTrackedRunOrigins(pending, paragraph.id, prepared.scope!, source);
      for (let step = 0; step < order.length; step++) {
        const index = order[step]!;
        // Adjacent marks share a Word address. Re-read the exact retained site after reopening.
        const sites = collectRevisionSites(pending).filter((site) => site.paragraphMark);
        expect(sites).toHaveLength(2 - step);
        const site = sites[step === 0 ? index : 0]!;
        const read = (name: string) =>
          site.node.attributes.find(
            (attribute) =>
              attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === name
          )?.value;
        const revision = {
          author: read('author')!,
          id: read('id')!,
          ...(read('date') ? { date: read('date') } : {}),
        };
        pending = apply(pending, {
          op: decisions[index] === 'accept' ? 'acceptRevision' : 'rejectRevision',
          revision,
          siteNodeIds: [site.node.id],
        });
        const opened = readOoxmlPart(serializeOoxmlPart(pending), {
          name: pending.name,
          contentType: pending.contentType,
        });
        if (!opened.ok) throw Error(opened.reason);
        pending = opened.part;
      }
      const retained = [1, 3].filter((_, index) => decisions[index] === 'accept');
      const direct = retained.length
        ? apply(source, { op: 'splitParagraphMany', paragraphId: id, offsets: retained })
        : source;
      expect(serializeOoxmlPart(pending)).toBe(serializeOoxmlPart(direct));
      expect(canonicalOoxmlFingerprint(pending)).toBe(canonicalOoxmlFingerprint(direct));
    });
