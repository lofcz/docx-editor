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
}
