import { expect, test } from 'bun:test';
import {
  nativeTextSliceOf,
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
