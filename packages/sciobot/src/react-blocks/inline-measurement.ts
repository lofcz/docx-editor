import {
  setSurfaceInlineWidgetMetrics,
  type PaginatedSurface,
  type SurfacePaintFrame,
} from '@docx-editor.dev/core/editor';
import type { InlineWidgetMetrics } from '@docx-editor.dev/core/layout';
import { inlineLayoutBoxes } from './inline-layout-boxes';

/** Intrinsic CSS pixel metrics at 100% zoom. Return null to retain native fallback. */
export type MeasureInlineElement = (
  element: HTMLElement,
  baseline: HTMLElement,
  zoom: number
) => InlineWidgetMetrics | null;

const measureElement: MeasureInlineElement = (element, baseline, zoom) => {
  const rect = element.getBoundingClientRect(),
    y = baseline.getBoundingClientRect().top;
  return {
    width: rect.width / zoom,
    ascent: (y - rect.top) / zoom,
    descent: (rect.bottom - y) / zoom,
  };
};
interface Entry {
  readonly element: HTMLElement;
  readonly baseline: HTMLElement;
  readonly zoom: number;
  readonly observer: ResizeObserver | null;
}

/** Local intrinsic measurements. Native layout remains the authority for placement. */
export class InlineMeasurement {
  private readonly entries = new Map<string, Entry>();
  private metrics = new Map<string, InlineWidgetMetrics>();
  private readonly hidden = new Map<HTMLElement, string>();
  private frame: SurfacePaintFrame | null = null;
  private queued = false;
  private disposed = false;
  private revision = 0;

  constructor(
    private readonly surface: PaginatedSurface,
    private readonly measure: MeasureInlineElement = measureElement
  ) {}

  track(controlId: string, element: HTMLElement, baseline: HTMLElement, zoom: number): void {
    if (this.disposed) return;
    const previous = this.entries.get(controlId);
    let observer = previous?.observer ?? null;
    if (previous?.element !== element) {
      observer?.disconnect();
      const Observer = element.ownerDocument.defaultView?.ResizeObserver;
      observer = Observer ? new Observer(() => this.schedule()) : null;
      observer?.observe(element);
    }
    this.entries.set(controlId, { element, baseline, zoom, observer });
    this.schedule();
  }

  release(controlId: string, element: HTMLElement): void {
    const entry = this.entries.get(controlId);
    if (entry?.element !== element) return;
    entry.observer?.disconnect();
    this.entries.delete(controlId);
    this.schedule();
  }

  schedule(): void {
    if (this.queued || this.disposed) return;
    this.queued = true;
    queueMicrotask(() => {
      this.queued = false;
      this.flush();
    });
  }

  private flush(): void {
    if (this.disposed) return;
    const next = new Map<string, InlineWidgetMetrics>();
    for (const [id, entry] of this.entries) {
      // Virtual pages park roots. Their last valid measurement must remain available.
      if (!entry.element.isConnected) {
        const old = this.metrics.get(id);
        if (old) next.set(id, old);
        continue;
      }
      const result = this.measure(entry.element, entry.baseline, entry.zoom);
      if (result) next.set(id, Object.freeze({ ...result }));
    }
    const equal =
      next.size === this.metrics.size &&
      [...next].every(([id, value]) => {
        const old = this.metrics.get(id);
        return (
          old &&
          Object.is(old.width, value.width) &&
          Object.is(old.ascent, value.ascent) &&
          Object.is(old.descent, value.descent)
        );
      });
    if (equal) {
      this.place();
      return;
    }
    this.metrics = next;
    setSurfaceInlineWidgetMetrics(
      this.surface,
      next.size
        ? {
            measure: ({ controlId }) => next.get(controlId) ?? null,
            revision: String(++this.revision),
          }
        : null
    );
    this.place();
  }

  place(frame = this.frame): void {
    if (this.disposed || !frame) return;
    this.frame = frame;
    this.restoreFallback();
    const native = new Map<string, HTMLElement>();
    for (const span of frame.pagesLayer.querySelectorAll<HTMLElement>(
      '[data-docx-inline-widget]'
    )) {
      const page = span.closest<HTMLElement>('.docx-page');
      native.set(`${page?.dataset.pageIndex}:${span.dataset.docxInlineWidget}`, span);
    }
    const boxes = inlineLayoutBoxes(frame.layout);
    for (const [id, { element }] of this.entries) {
      const root = element.closest<HTMLElement>('[data-sciobot-block-host]');
      if (!root) continue;
      const key = `${root.closest<HTMLElement>('.docx-page')?.dataset.pageIndex}:${id}`;
      const box = boxes.get(key),
        fallback = native.get(key);
      root.style.visibility = box && fallback ? 'visible' : 'hidden';
      root.style.overflow = 'hidden';
      if (!box || !fallback) continue;
      Object.assign(root.style, {
        left: `${box.x * frame.scale}px`,
        top: `${box.y * frame.scale}px`,
        width: `${box.width * frame.scale}px`,
        height: `${box.height * frame.scale}px`,
      });
      this.hidden.set(fallback, fallback.style.visibility);
      fallback.style.visibility = 'hidden';
    }
  }

  private restoreFallback(): void {
    for (const [element, visibility] of this.hidden) element.style.visibility = visibility;
    this.hidden.clear();
  }

  beforePaint(): void {
    this.restoreFallback();
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    for (const entry of this.entries.values()) entry.observer?.disconnect();
    this.entries.clear();
    this.metrics.clear();
    this.restoreFallback();
    this.frame = null;
    setSurfaceInlineWidgetMetrics(this.surface, null);
  }
}
