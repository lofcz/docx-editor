import { applyTreeOp, bodyStoryRoot, paragraphTextOf, storyParagraphs } from '../index.ts';
import { inlineTextMerged } from '../store/tree-op-inline-slices.ts';
import { readOoxmlPart, type OoxmlNode } from '../package/ooxml-tree.ts';
import { expect, test } from 'bun:test';
import {
  nativeTextSliceOf,
  nativeTextSpliceOf,
  observeCanonicalPrimitiveJournal,
  recordNativeFreshText,
  recordNativeTextMerge,
  recordNativeTextSlice,
  recordPutNode,
  recordSpliceText,
  runObservedStoreTransaction,
} from '../package/canonical-primitive-capture.ts';
import type { CanonicalPrimitiveJournal } from '../package/canonical-primitive-journal.ts';

test('a coalesced native leaf retains ordered old windows and authored characters through another split', () => {
  const store = {};
  const journals: CanonicalPrimitiveJournal[] = [];
  const stop = observeCanonicalPrimitiveJournal(store, (journal) => journals.push(journal));
  runObservedStoreTransaction(
    store,
    () => {
      const original = { id: 'original', value: 'alpha beta' };
      const first = { id: 'first', value: 'alpha' };
      const fresh = { id: 'fresh', value: ' NEW ' };
      const last = { id: 'last', value: 'beta' };
      recordNativeTextSlice(original, first, 0);
      recordNativeFreshText(fresh);
      recordSpliceText('unrelated', 0, 0, '!');
      recordNativeTextSlice(original, last, 6);
      const joined = { id: 'joined', value: 'alpha NEW ' };
      recordNativeTextMerge(first, fresh, joined);
      const complete = { id: 'complete', value: 'alpha NEW beta' };
      recordNativeTextMerge(joined, last, complete);
      const cut = { id: 'cut', value: 'ha NEW be' };
      recordNativeTextSlice(complete, cut, 3);
      recordPutNode({ logicalId: cut.id, kind: 'textValue', sourceText: nativeTextSliceOf(cut) });
    },
    () => true
  );
  stop();
  const descriptor = journals[0]!.effects[1];
  expect(descriptor?.kind).toBe('putNode');
  if (descriptor?.kind !== 'putNode' || descriptor.descriptor.kind !== 'textValue')
    throw Error('Missing leaf');
  expect(descriptor.descriptor.sourceText).toEqual([
    { logicalId: 'original', utf16Start: 3, utf16Length: 2, beforeEffect: 0 },
    { utf16Length: 5 },
    { logicalId: 'original', utf16Start: 6, utf16Length: 2, beforeEffect: 1 },
  ]);
  expect(Object.isFrozen(descriptor.descriptor.sourceText)).toBe(true);
  for (const piece of descriptor.descriptor.sourceText!) expect(Object.isFrozen(piece)).toBe(true);
  expect(structuredClone(journals[0])).toEqual(journals[0]);
});

test('a retained native tail lends its exact pre-edit window to later clones', () => {
  const store = {};
  const stop = observeCanonicalPrimitiveJournal(store, () => {});
  runObservedStoreTransaction(
    store,
    () => {
      const original = { id: 'original', value: 'left middle right' };
      const tail = { id: 'original', value: 'middle right' };
      recordNativeTextSlice(original, tail, 5);
      recordSpliceText('original', 0, 5, '');
      const middle = { id: 'middle', value: 'middle' };
      recordNativeTextSlice(tail, middle, 0);
      expect(nativeTextSliceOf(middle)).toEqual([
        { logicalId: 'original', utf16Start: 5, utf16Length: 6, beforeEffect: 0 },
      ]);
    },
    () => true
  );
  stop();
});

test('coalescing an empty imported text element creates no nonexistent source window', () => {
  const part = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t/><w:t>Word</w:t></w:r></w:p></w:body></w:document>',
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!part.ok) throw Error('Invalid fixture');
  const leaves: OoxmlNode[] = [];
  const walk = (node: OoxmlNode) => {
    if (node.kind === 'text') leaves.push(node);
    if (node.kind !== 'textValue') node.children.forEach(walk);
  };
  walk(part.part.root);
  const store = {};
  const stop = observeCanonicalPrimitiveJournal(store, () => {});
  let next = 0;
  runObservedStoreTransaction(
    store,
    () => {
      for (const [left, right] of [
        [leaves[0]!, leaves[1]!],
        [leaves[1]!, leaves[0]!],
      ]) {
        const merged = inlineTextMerged(() => `new-${next++}`, left!, right!);
        if (merged.kind === 'textValue' || merged.children[0]?.kind !== 'textValue')
          throw Error('Missing text');
        expect(merged.children[0].value).toBe('Word');
        const original = leaves[1]!;
        if (original.kind === 'textValue') throw Error('Missing element');
        expect(nativeTextSliceOf(merged.children[0])).toEqual([
          { logicalId: original.children[0]!.id, utf16Start: 0, utf16Length: 4, beforeEffect: 0 },
        ]);
      }
    },
    () => true
  );
  stop();
});

test('coalescing native text preserves boundary spaces when imported xml:space is default', () => {
  const part = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve"> left </w:t><w:t xml:space="default"> right </w:t></w:r></w:p></w:body></w:document>',
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!part.ok) throw Error('Invalid fixture');
  const texts: OoxmlNode[] = [];
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'text') texts.push(node);
    if (node.kind !== 'textValue') node.children.forEach(visit);
  };
  visit(part.part.root);
  let next = 0;
  const merged = inlineTextMerged(() => `joined-${next++}`, texts[0]!, texts[1]!);
  if (merged.kind === 'textValue' || merged.children[0]?.kind !== 'textValue')
    throw Error('Missing text');
  // The reader drops the boundary spaces that Word does not show without xml:space="preserve".
  expect(merged.children[0].value).toBe(' left right');
  expect(
    merged.attributes.some((attr) => attr.localName === 'space' && attr.value === 'default')
  ).toBe(false);
});

for (const [value, offset, insert] of [
  ['A ', 1, ' peer'],
  ['aaaa', 1, 'a'],
  ['😀 😀', 3, '😀 '],
] as const)
  test(`native typing keeps its exact primitive boundary: ${JSON.stringify([value, offset, insert])}`, () => {
    const loaded = readOoxmlPart(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p></w:body></w:document>`,
      { name: '/word/document.xml', contentType: 'application/xml' }
    );
    if (!loaded.ok) throw Error(loaded.reason);
    const store = {},
      journals: CanonicalPrimitiveJournal[] = [];
    const stop = observeCanonicalPrimitiveJournal(store, (journal) => journals.push(journal));
    const paragraphId = storyParagraphs(bodyStoryRoot(loaded.part)!)[0]!.id;
    let output = loaded.part;
    runObservedStoreTransaction(
      store,
      () => {
        const result = applyTreeOp(loaded.part, {
          op: 'insertText',
          paragraphId,
          offset,
          text: insert,
        });
        if (!result.ok) throw Error(result.reason);
        output = result.part;
      },
      () => true
    );
    stop();
    expect(paragraphTextOf(output, paragraphId)).toBe(
      value.slice(0, offset) + insert + value.slice(offset)
    );
    const splices = journals[0]!.effects.filter((e) => e.kind === 'spliceText');
    expect(splices).toHaveLength(1);
    expect(splices[0]).toMatchObject({ utf16Start: offset, deleteCount: 0, insert });
    expect(Object.keys(splices[0]!).sort()).toEqual([
      'deleteCount',
      'insert',
      'kind',
      'logicalId',
      'utf16Start',
    ]);
    const leaf = (node: OoxmlNode): OoxmlNode | null =>
      node.kind === 'textValue' ? node : (node.children.map(leaf).find(Boolean) ?? null);
    expect(nativeTextSpliceOf(leaf(loaded.part.root)!, leaf(output.root)!)).toBeUndefined();
  });
