import { expect, test } from 'bun:test';
import { furnitureWaits } from '../../layout/published-story-projections.ts';
import type { LayoutSteps } from '../../layout/layout-steps.ts';
import type { SemanticLayout } from '../../layout/semantic-records.ts';
import { createDeferredLayout } from '../surface-deferred-layout.ts';
import { furnitureLaterSteps } from '../surface-furniture-later.ts';

const layout: SemanticLayout = { revision: 1, pages: [] };

test('a short field-free pass completes without either viewport lookup', () => {
  let reads = 0;
  const deferred = createDeferredLayout({
    now: () => 100,
    pageAfterView: () => ++reads,
    current: () => layout,
    revision: () => 1,
    publish: () => {},
    snapshot: () => () => {},
  });
  function* steps(): LayoutSteps<SemanticLayout> {
    yield { pages: [], finalize: () => layout };
    return layout;
  }
  const result = deferred.run(
    furnitureLaterSteps(steps(), () => new Set([++reads])),
    1
  );
  expect(result).toBe(layout);
  expect(reads).toBe(0);
  expect(deferred.pending()).toBe(false);
});

test('fields capture the visible pages once for both prefix and final projections', () => {
  let reads = 0;
  const visible = new Set([4, 5]);
  function* steps(): LayoutSteps<SemanticLayout> {
    yield {
      pages: [],
      finalize() {
        expect(furnitureWaits(1)).toBe(true);
        expect(furnitureWaits(2)).toBe(false);
        return layout;
      },
    };
    expect(furnitureWaits(7)).toBe(false);
    expect(furnitureWaits(8)).toBe(true);
    return layout;
  }
  const work = furnitureLaterSteps(steps(), () => {
    reads++;
    return visible;
  });
  const prefix = work.next();
  expect(reads).toBe(0);
  if (prefix.done || !prefix.value?.finalize) throw new Error('Missing layout prefix');
  expect(prefix.value.finalize([])).toBe(layout);
  visible.clear();
  expect(work.next()).toEqual({ done: true, value: layout });
  expect(reads).toBe(1);
  expect(furnitureWaits(8)).toBe(false);
});

test('a hidden editor projects every field and cancelling it releases the predicate', () => {
  let reads = 0;
  function* steps(): LayoutSteps<SemanticLayout> {
    expect(furnitureWaits(0)).toBe(false);
    expect(furnitureWaits(20)).toBe(false);
    yield;
    return layout;
  }
  const work = furnitureLaterSteps(steps(), () => {
    reads++;
    return undefined;
  });
  expect(work.next().done).toBe(false);
  expect(work.return(layout).value).toBe(layout);
  expect(furnitureWaits(20)).toBe(false);
  expect(reads).toBe(1);
});
