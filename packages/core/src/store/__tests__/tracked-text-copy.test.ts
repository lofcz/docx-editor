import { expect, test } from 'bun:test';
import {
  TreeDocumentStore,
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  withPart,
} from '../index.ts';
import { findNode } from '../package/ooxml-edit.ts';
import { canonicalOoxmlFingerprint } from '../package/ooxml-serialize.ts';
import { readOoxmlPart, serializeOoxmlPart, type OoxmlNode } from '../package/ooxml-tree.ts';
import {
  flushPendingCanonicalJournals,
  runObservedStoreTransaction,
  observeCanonicalPrimitiveJournal,
  nativeTextSliceOf,
} from '../package/canonical-primitive-capture.ts';
import type { CanonicalPrimitiveJournal } from '../package/canonical-primitive-journal.ts';
import { copyTrackedText } from '../store/tracked-text-copy.ts';
import { ADA, apply, paragraphId, part } from './tracked-edit-fixture.ts';
const leaves = (node: OoxmlNode): Extract<OoxmlNode, { kind: 'textValue' }>[] =>
  node.kind === 'textValue' ? [node] : node.children.flatMap(leaves);
for (const styled of [false, true])
  for (const [value, start, end] of [
    ['Original', 0, 8],
    ['word word word', 5, 9],
    ['A𝛼éZ', 1, 5],
  ] as const)
    test(`tracked hyperlink copies carry exact source recipes ${value}/${start}/${end}/${styled}`, () => {
      const before = part(
        `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`
      );
      const source = leaves(before.root)[0]!;
      const prepared = prepareTrackedRunOrigins(before, paragraphId(before), [{ start, end }]);
      const store = new TreeDocumentStore(prepared.part);
      const captured: CanonicalPrimitiveJournal[] = [];
      const stop = observeCanonicalPrimitiveJournal(store, (journal) => captured.push(journal));
      try {
        const operation = {
          op: 'insertHyperlink',
          paragraphId: paragraphId(before),
          start,
          end,
          anchor: 'target',
          ...(styled ? { styleId: 'Hyperlink' } : {}),
          revision: ADA,
        } as const;
        const result = runObservedStoreTransaction(
          store,
          () =>
            store.transact((ctx) => {
              if (!ctx.apply(operation)) throw Error('Tracked insertion refused.');
              if (prepared.scope)
                ctx.applyPackage((pkg) =>
                  withPart(
                    pkg,
                    sealTrackedRunOrigins(
                      pkg.parts.get(before.name)!,
                      paragraphId(before),
                      prepared.scope!,
                      before
                    )
                  )
                );
            }),
          (value) => value.ok
        );
        if (!result.ok) throw Error(result.reason);
        flushPendingCanonicalJournals(store);
        expect(captured).toHaveLength(1);
        const copies = captured[0]!.effects.flatMap((effect, index) => {
          if (effect.kind !== 'putNode' || effect.descriptor.kind !== 'textValue') return [];
          const node = findNode(store.part, effect.descriptor.logicalId);
          return node?.kind === 'textValue' && node.value === value.slice(start, end)
            ? [{ descriptor: effect.descriptor, index }]
            : [];
        });
        expect(copies).toHaveLength(2);
        for (const { descriptor, index } of copies) {
          expect(descriptor.sourceText).toHaveLength(1);
          const recipe = descriptor.sourceText![0]!;
          expect(recipe).toMatchObject({
            logicalId: source.id,
            utf16Start: start,
            utf16Length: end - start,
          });
          if (!('beforeEffect' in recipe)) throw Error('Missing native source time.');
          expect(recipe.beforeEffect).toBeLessThanOrEqual(index);
          expect(recipe.beforeEffect).toBeGreaterThanOrEqual(0);
        }
        const pending = store.part;
        for (const decision of ['accept', 'reject'] as const) {
          const decided = apply(pending, {
            op: decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
          });
          const expected =
            decision === 'accept' ? apply(before, { ...operation, revision: undefined }) : before;
          expect(canonicalOoxmlFingerprint(decided)).toBe(canonicalOoxmlFingerprint(expected));
          const reopened = readOoxmlPart(serializeOoxmlPart(decided), {
            name: decided.name,
            contentType: decided.contentType,
          });
          if (!reopened.ok) throw Error(reopened.reason);
          expect(canonicalOoxmlFingerprint(reopened.part)).toBe(
            canonicalOoxmlFingerprint(expected)
          );
        }
      } finally {
        stop();
      }
    });

test('copied native text records positional concatenation across distinct source leaves', () => {
  const source = part('<w:p><w:r><w:t>same</w:t></w:r><w:r><w:t>same</w:t></w:r></w:p>');
  const projected = part('<w:p><w:r><w:t>samesame</w:t></w:r></w:p>');
  const host = {};
  const stop = observeCanonicalPrimitiveJournal(host, () => {});
  let counter = 0;
  try {
    runObservedStoreTransaction(
      host,
      () => {
        const copied = copyTrackedText([projected.root], [source.root], () => `copy-${counter++}`);
        const target = leaves(copied[0]!)[0]!;
        const recipe = nativeTextSliceOf(target)!;
        const original = leaves(source.root);
        expect(recipe).toEqual(
          original.map((leaf) => ({
            logicalId: leaf.id,
            utf16Start: 0,
            utf16Length: 4,
            beforeEffect: 0,
          }))
        );
        expect(original[0]!.id).not.toBe(original[1]!.id);
      },
      () => false
    );
  } finally {
    stop();
  }
});

test('tracked hyperlink copies can opt out of source lineage for concurrent source deletion', () => {
  const source = part('<w:p><w:r><w:t>site</w:t></w:r></w:p>');
  const copied = copyTrackedText([source.root], [source.root], () => 'copy', false);
  expect(nativeTextSliceOf(leaves(copied[0]!)[0]!)).toBeUndefined();
});

for (const target of ['changed', 'same extra'])
  test(`native text copy refuses an inexact source projection: ${target}`, () => {
    const source = part('<w:p><w:r><w:t>same</w:t></w:r></w:p>');
    const projected = part(`<w:p><w:r><w:t>${target}</w:t></w:r></w:p>`);
    const host = {},
      stop = observeCanonicalPrimitiveJournal(host, () => {});
    let counter = 0;
    try {
      expect(() =>
        runObservedStoreTransaction(
          host,
          () => copyTrackedText([projected.root], [source.root], () => `copy-${counter++}`),
          () => false
        )
      ).toThrow();
    } finally {
      stop();
    }
  });
