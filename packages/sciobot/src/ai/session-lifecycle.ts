export type ReviewSessionDocument = string;

export interface ReviewSessionHandle {
  readonly documentId: ReviewSessionDocument;
  readonly generation: number;
  readonly ended: boolean;
  /** Queue work for the current session. Stale work is discarded before it runs. */
  queue(run: () => void): boolean;
  /** Cancel queued work without ending the session. */
  cancelQueued(): void;
  /** Finish once. `flush` runs queued work before ending; `cancel` drops it. */
  end(mode?: 'flush' | 'cancel'): boolean;
  /** True only while this handle is the current document session. */
  isCurrent(): boolean;
}

export interface ReviewSessionLifecycleOptions {
  readonly schedule?: (run: () => void) => () => void;
}

interface Pending {
  readonly run: () => void;
  cancel: () => void;
  active: boolean;
}

/**
 * Owns delayed review callbacks by document identity and generation.
 * A callback checks its owner again when it runs, so a stale frame cannot edit a new document.
 */
export function createReviewSessionLifecycle(options: ReviewSessionLifecycleOptions = {}): {
  begin(documentId: ReviewSessionDocument): ReviewSessionHandle;
  current(): ReviewSessionHandle | null;
} {
  const schedule =
    options.schedule ??
    ((run: () => void) => {
      if (typeof requestAnimationFrame === 'function') {
        const frame = requestAnimationFrame(run);
        return () => cancelAnimationFrame(frame);
      }
      const timer = setTimeout(run, 0);
      return () => clearTimeout(timer);
    });
  let generation = 0;
  let active: Session | null = null;

  class Session implements ReviewSessionHandle {
    readonly generation = ++generation;
    readonly pending = new Set<Pending>();
    ended = false;
    constructor(readonly documentId: ReviewSessionDocument) {}

    isCurrent() {
      return !this.ended && active === this;
    }

    queue(run: () => void) {
      if (typeof run !== 'function' || !this.isCurrent()) return false;
      const pending: Pending = {
        run,
        cancel: () => {},
        active: true,
      };
      this.pending.add(pending);
      pending.cancel = schedule(() => {
        if (!pending.active) return;
        pending.active = false;
        this.pending.delete(pending);
        if (this.isCurrent()) pending.run();
      });
      return true;
    }

    cancelQueued() {
      for (const pending of this.pending) {
        pending.active = false;
        pending.cancel();
      }
      this.pending.clear();
    }

    end(mode: 'flush' | 'cancel' = 'cancel') {
      if (this.ended) return false;
      if (mode === 'flush') {
        const pending = [...this.pending];
        for (const item of pending) {
          if (!item.active) continue;
          item.active = false;
          item.cancel();
          this.pending.delete(item);
          if (this.isCurrent()) item.run();
        }
      } else this.cancelQueued();
      this.ended = true;
      if (active === this) active = null;
      return true;
    }
  }

  return {
    begin(documentId) {
      if (typeof documentId !== 'string' || !documentId)
        throw new TypeError('Invalid document ID.');
      active?.end('cancel');
      active = new Session(documentId);
      return active;
    },
    current: () => active,
  };
}
