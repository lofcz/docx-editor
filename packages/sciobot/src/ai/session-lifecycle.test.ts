import { expect, test } from 'bun:test';
import { createReviewSessionLifecycle } from './session-lifecycle';

function scheduler() {
  const queue: (() => void)[] = [];
  return {
    schedule(run: () => void) {
      queue.push(run);
      return () => {
        const index = queue.indexOf(run);
        if (index >= 0) queue.splice(index, 1);
      };
    },
    flush() {
      queue.splice(0).forEach((run) => run());
    },
  };
}

test('a document switch cancels stale callbacks before they can write', () => {
  const clock = scheduler(),
    lifecycle = createReviewSessionLifecycle({ schedule: clock.schedule });
  const first = lifecycle.begin('one'),
    calls: string[] = [];
  expect(first.queue(() => calls.push('stale'))).toBe(true);
  const second = lifecycle.begin('two');
  expect(first.ended).toBe(true);
  expect(first.isCurrent()).toBe(false);
  expect(second.isCurrent()).toBe(true);
  clock.flush();
  expect(calls).toEqual([]);
});

test('ending with flush runs each queued callback once and blocks later work', () => {
  const clock = scheduler(),
    lifecycle = createReviewSessionLifecycle({ schedule: clock.schedule });
  const session = lifecycle.begin('doc'),
    calls: string[] = [];
  expect(session.queue(() => calls.push('one'))).toBe(true);
  expect(session.queue(() => calls.push('two'))).toBe(true);
  expect(session.end('flush')).toBe(true);
  expect(calls).toEqual(['one', 'two']);
  clock.flush();
  expect(calls).toEqual(['one', 'two']);
  expect(session.queue(() => calls.push('late'))).toBe(false);
  expect(session.end('flush')).toBe(false);
});

test('cancelQueued preserves the session while dropping delayed review work', () => {
  const clock = scheduler(),
    lifecycle = createReviewSessionLifecycle({ schedule: clock.schedule });
  const session = lifecycle.begin('doc'),
    calls: string[] = [];
  session.queue(() => calls.push('discard'));
  session.cancelQueued();
  clock.flush();
  expect(calls).toEqual([]);
  expect(session.isCurrent()).toBe(true);
  expect(session.queue(() => calls.push('kept'))).toBe(true);
  clock.flush();
  expect(calls).toEqual(['kept']);
});
