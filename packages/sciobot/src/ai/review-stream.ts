import { createReviewSessionLifecycle, type ReviewSessionHandle } from './session-lifecycle';

export interface ReviewStreamApplyContext {
  readonly documentId: string;
  readonly generation: number;
  readonly history: 'start' | 'continue';
  readonly final: boolean;
}

export interface ReviewStreamOptions<T, R> {
  readonly documentId: string;
  readonly apply: (target: T, context: ReviewStreamApplyContext) => R;
  readonly schedule?: (run: () => void) => () => void;
  readonly onError?: (error: unknown) => void;
}

export interface ReviewStream<T, R> {
  readonly session: ReviewSessionHandle;
  readonly state: 'active' | 'ended';
  /** Keep only the latest complete target until the next animation frame. */
  push(target: T): boolean;
  /** Apply the currently buffered target immediately. */
  flush(): R | undefined;
  /** Apply the final target exactly once, then end the session. */
  finish(target?: T): R | undefined;
  /** Drop buffered work and end without changing the document. */
  cancel(): boolean;
}

/** Coalesces streamed review targets while preserving one history batch per session. */
export function beginReviewStream<T, R>(options: ReviewStreamOptions<T, R>): ReviewStream<T, R> {
  const lifecycle = createReviewSessionLifecycle({ schedule: options.schedule });
  const session = lifecycle.begin(options.documentId);
  let latest: T | undefined;
  let hasLatest = false;
  let frameQueued = false;
  let state: ReviewStream<T, R>['state'] = 'active';
  let started = false;
  let result: R | undefined;

  const applyLatest = (final: boolean) => {
    if (!hasLatest || state !== 'active') return undefined;
    const target = latest as T;
    latest = undefined;
    hasLatest = false;
    frameQueued = false;
    try {
      result = options.apply(target, {
        documentId: session.documentId,
        generation: session.generation,
        history: started ? 'continue' : 'start',
        final,
      });
      started = true;
      return result;
    } catch (error) {
      options.onError?.(error);
      return undefined;
    }
  };
  const stream: ReviewStream<T, R> = {
    session,
    get state() {
      return state;
    },
    push(target) {
      if (state !== 'active') return false;
      latest = target;
      hasLatest = true;
      if (!frameQueued) {
        frameQueued = true;
        session.queue(() => applyLatest(false));
      }
      return true;
    },
    flush() {
      if (state !== 'active') return undefined;
      session.cancelQueued();
      return applyLatest(false);
    },
    finish(target) {
      if (state !== 'active') return undefined;
      if (target !== undefined) {
        latest = target;
        hasLatest = true;
      }
      session.cancelQueued();
      const applied = applyLatest(true);
      state = 'ended';
      session.end('cancel');
      return applied;
    },
    cancel() {
      if (state !== 'active') return false;
      latest = undefined;
      hasLatest = false;
      frameQueued = false;
      state = 'ended';
      return session.end('cancel');
    },
  };
  return stream;
}
