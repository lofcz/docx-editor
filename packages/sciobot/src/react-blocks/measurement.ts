import type { PaginatedSurface } from '@docx-editor.dev/core/editor';
import { contentControlContentNodeOf, type OoxmlNode } from '@docx-editor.dev/core/store';

/** Widget height in CSS pixels at the current zoom. */
export type MeasureBlockElement = (element: HTMLElement) => number;

/** What one measured root needs to turn its height into reserved native flow. */
export interface MeasuredFragment {
  /** The block's last body paragraph: the reservation lands after its last line. */
  readonly paragraphId: string;
  /** The complete native fallback height, in points. It never includes a reservation. */
  readonly nativeHeight: number;
  readonly scale: number;
  /** Screen pixels per CSS pixel in the widget. Intrinsic overflow keeps its full reservation. */
  readonly contentScale?: number;
  /** Reservation budget in screen pixels. Overflow beyond this limit does not reserve flow. */
  readonly maxHeight?: number;
}

function defaultMeasure(element: HTMLElement, fragment: MeasuredFragment): number {
  const viewportHeight = element.getBoundingClientRect().height;
  return fragment.contentScale === undefined
    ? viewportHeight
    : Math.max(viewportHeight, element.scrollHeight * fragment.contentScale);
}

/**
 * Turns measured widget heights into layout-only native reservations.
 *
 * The reservation is a surface view input, so user history, collaboration peers, and saved
 * files never see a measurement. Native fragment heights exclude reservations. Intrinsic
 * widget content keeps the same reservation budget when page placement shrinks its viewport,
 * so a reservation cannot feed back into the next measurement.
 */
export class BlockMeasurement {
  private readonly tracked = new Map<
    string,
    { element: HTMLElement; fragment: MeasuredFragment; height?: number }
  >();
  private readonly observer: ResizeObserver | null;
  private queued = false;
  private disposed = false;
  private published: ReadonlyMap<string, number> = new Map();

  constructor(
    private readonly surface: PaginatedSurface,
    private readonly measure?: MeasureBlockElement
  ) {
    this.observer =
      typeof ResizeObserver === 'function' ? new ResizeObserver(() => this.schedule()) : null;
  }

  /** Track one root's measuring element. Replaces an earlier element for the same key. */
  track(key: string, element: HTMLElement, fragment: MeasuredFragment): void {
    if (this.disposed) return;
    const previous = this.tracked.get(key);
    if (previous?.element !== element) {
      if (previous) this.observer?.unobserve(previous.element);
      this.observer?.observe(element);
    }
    this.tracked.set(key, { element, fragment, height: previous?.height });
    this.schedule();
  }

  release(key: string): void {
    const previous = this.tracked.get(key);
    if (!previous) return;
    this.observer?.unobserve(previous.element);
    this.tracked.delete(key);
    this.schedule();
  }

  /** Coalesce measurements from one paint or resize burst into one reservation update. */
  schedule(): void {
    if (this.queued || this.disposed) return;
    this.queued = true;
    queueMicrotask(() => {
      this.queued = false;
      this.flush();
    });
  }

  /** Measure every tracked root now and publish the reservations. */
  flush(): void {
    if (this.disposed) return;
    const reservations = new Map<string, number>();
    for (const entry of this.tracked.values()) {
      const { element, fragment } = entry;
      // A root parked with its virtual page keeps the height it last measured. Dropping its
      // room moved the content, which brought that page back and measured it again, in a loop.
      if (element.isConnected && !element.closest('[hidden]')) {
        const height = this.measure ? this.measure(element) : defaultMeasure(element, fragment);
        if (Number.isFinite(height)) entry.height = height;
      }
      if (entry.height === undefined || !(fragment.scale > 0)) continue;
      const extra =
        Math.min(entry.height, fragment.maxHeight ?? Infinity) / fragment.scale -
        fragment.nativeHeight;
      if (extra > 0)
        reservations.set(
          fragment.paragraphId,
          Math.max(reservations.get(fragment.paragraphId) ?? 0, extra)
        );
    }
    this.published = reservations;
    this.surface.setBlockReservations(reservations);
  }

  /** The room last reserved after one paragraph, in points, or 0. */
  reservedAfter(paragraphId: string): number {
    return this.published.get(paragraphId) ?? 0;
  }

  dispose(): void {
    if (this.disposed) return;
    this.observer?.disconnect();
    this.tracked.clear();
    this.published = new Map();
    this.surface.setBlockReservations(new Map());
    this.disposed = true;
  }
}

/**
 * The control's last body paragraph, or null when it ends in a table. Native reservations
 * apply to body flow; cell paragraphs keep their table geometry.
 */
export function lastBodyParagraphIn(control: OoxmlNode): string | null {
  const content = contentControlContentNodeOf(control);
  if (!content) return null;
  for (let index = content.children.length - 1; index >= 0; index -= 1) {
    const child = content.children[index]!;
    if (child.kind === 'paragraph') return child.id;
    if (child.kind === 'contentControl') return lastBodyParagraphIn(child);
    if (child.kind === 'table') return null;
  }
  return null;
}
