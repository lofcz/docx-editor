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
  beforePaint?(): void;
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
