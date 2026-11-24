// Whether the browser holds input the surface has not handled yet (Chromium's
// `navigator.scheduling.isInputPending`). Every other engine answers false, which keeps the
// surface on its synchronous paths.

import type { HostTimers } from './host-realm.ts';
import type { FieldResultsScope } from './field-results-scope.ts';

/** Own the paint deadline for one surface's buffered plain text. */
export function createTypeFlushTimer(
  timers: HostTimers,
  scope: FieldResultsScope,
  container: () => Element,
  flush: () => void,
  continuesTyping: () => boolean
) {
  let timer: number | null = null;
  function schedule(delay: number): void {
    timer = timers.setTimeout(
      scope.wrap(() => {
        timer = null;
        flush();
      }),
      delay
    );
  }
  function cancel(): void {
    if (timer === null) return;
    timers.clearTimeout(timer);
    timer = null;
  }
  return {
    current: () => timer,
    cancel,
    enqueue(): void {
      if (timer === null) schedule(0);
    },
    afterKey(inputType: string, previous: number | null): void {
      const started = timer;
      if (inputType !== 'insertText' || previous !== null || started === null) return;
      queueMicrotask(
        scope.wrap(() => {
          if (timer !== started || pendingBrowserInput(container(), false) !== false) return;
          if (continuesTyping()) {
            // Keep a rapid continuation for half a 60 Hz frame. Further characters
            // use this same timer. A continuous stream cannot extend the deadline.
            cancel();
            schedule(9);
          } else {
            flush();
          }
        })
      );
    },
  };
}

/**
 * `includeContinuous` folds mousemove/wheel into the answer. Paint deferral wants that
 * (any input beats a repaint); the commit-tail LAYOUT deferral must not — a moving
 * pointer over a large document would otherwise defer every toolbar op, paste and
 * programmatic write, so that gate asks for discrete input (keys, clicks) only.
 */
export function browserInputPending(container: Element, includeContinuous = true): boolean {
  return pendingBrowserInput(container, includeContinuous) ?? false;
}

/** An unavailable input queue cannot confirm that a buffered key is the last key. */
export function pendingBrowserInput(
  container: Element,
  includeContinuous: boolean
): boolean | undefined {
  const scheduling = (
    container.ownerDocument.defaultView?.navigator as
      | (Navigator & {
          scheduling?: {
            isInputPending?: (options?: { includeContinuous?: boolean }) => boolean;
          };
        })
      | undefined
  )?.scheduling;
  return scheduling?.isInputPending?.({ includeContinuous });
}
