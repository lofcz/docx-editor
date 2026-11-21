import type { PreserveParagraphRoots } from '../output/semantic-paint-paragraph-roots.ts';
export type { PreserveParagraphRoots } from '../output/semantic-paint-paragraph-roots.ts';
import type { OoxmlPart } from '../store/package/ooxml-tree.ts';
import type { SemanticLayout } from '../layout/semantic-records.ts';

/** A completed page paint, before the editor restores its DOM selection. @public */
export interface SurfacePaintFrame {
  readonly pagesLayer: HTMLElement;
  readonly layout: SemanticLayout;
  readonly part: OoxmlPart;
  readonly scale: number;
}

/**
 * A host that borrows painted elements for custom document components.
 * Restore borrowed elements in beforePaint. Attach them in afterPaint, so the
 * editor can restore selection against their final position. Never mutate text.
 * @public
 */
export interface SurfacePaintObserver {
  /** Use the last painted frame on registration when this host has no native nodes to restore yet. */
  readonly replayCurrentFrame?: boolean;
  /** Paint keeps native nodes on these materialized pages. Restore all borrowed nodes when the set is absent. */
  beforePaint?(
    retainedPages?: ReadonlySet<number>,
    /** Opt in per page. The supplied roots keep their identity, but paint replaces their children.
     * Return false from the predicate to restore the host's loans normally. Valid only during this call. */
    preserveParagraphRoots?: PreserveParagraphRoots
  ): void;
  afterPaint(frame: SurfacePaintFrame): void;
  /** Release host resources when this surface is destroyed or its document is replaced. */
  destroyed?(): void;
}

/** Release every observer even when one host's cleanup fails. */
export function destroyPaintObservers(observers: Set<SurfacePaintObserver>): void {
  const current = [...observers];
  observers.clear();
  for (const observer of current) {
    try {
      observer.beforePaint?.();
    } catch (error) {
      console.error('Surface observer restoration failed.', error);
    }
    try {
      observer.destroyed?.();
    } catch (error) {
      console.error('Surface observer cleanup failed.', error);
    }
  }
}

/** Reuse the completed frame for observers that can join an existing paint. */
export function createPaintObservers(repaint: () => void, parked: () => boolean) {
  const observers = new Set<SurfacePaintObserver>();
  let lastFrame: SurfacePaintFrame | null = null;
  function beforePaint(retainedPages: ReadonlySet<number>, preserve: PreserveParagraphRoots): void {
    for (const observer of observers) observer.beforePaint?.(retainedPages, preserve);
  }
  return {
    subscribe(observer: SurfacePaintObserver) {
      observers.add(observer);
      if (lastFrame && (!observer.beforePaint || observer.replayCurrentFrame) && !parked())
        observer.afterPaint(lastFrame);
      else if (observer.beforePaint) repaint();
      return () => {
        if (observers.delete(observer)) observer.beforePaint?.();
      };
    },
    preparePaint() {
      lastFrame = null;
      // Only compute retained pages when a host borrows painted elements.
      for (const observer of observers) if (observer.beforePaint) return beforePaint;
      return undefined;
    },
    afterPaint(frame: SurfacePaintFrame) {
      lastFrame = frame;
      for (const observer of [...observers]) {
        if (observers.has(observer)) observer.afterPaint(frame);
      }
    },
    destroy() {
      lastFrame = null;
      destroyPaintObservers(observers);
    },
  };
}
