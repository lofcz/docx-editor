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
  private readonly observers = new Map<Document, ResizeObserver>();
  private readonly ids = new WeakMap<Element, string>();
  private readonly dirty = new Set<string>();
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
    const doc = element.ownerDocument;
    const view = doc.defaultView;
    const Observer = view?.ResizeObserver;
    let observer = this.observers.get(doc) ?? null;
    if (!observer && Observer) {
      // Chromium runs microtasks between observers. One delivery must publish one
      // layout snapshot when several equations acquire their intrinsic sizes together.
      observer = new Observer((changes) => {
        // Publish new dimensions after the browser completes this resize delivery.
        view.requestAnimationFrame(() => {
          for (const { target } of changes) {
            const id = this.ids.get(target);
            if (id !== undefined) this.schedule(id);
          }
        });
      });
      this.observers.set(doc, observer);
    }
    if (previous?.element !== element || previous.observer !== observer) {
      if (previous) {
        previous.observer?.unobserve(previous.element);
        this.ids.delete(previous.element);
      }
      this.ids.set(element, controlId);
      observer?.observe(element);
    }
    this.entries.set(controlId, { element, baseline, zoom, observer });
    this.schedule(controlId);
  }

  release(controlId: string, element: HTMLElement): void {
    const entry = this.entries.get(controlId);
    if (entry?.element !== element) return;
    entry.observer?.unobserve(element);
    this.ids.delete(element);
    this.entries.delete(controlId);
    this.schedule(controlId);
  }

  remeasure(): void {
    for (const id of this.entries.keys()) this.schedule(id);
  }

  private schedule(id: string): void {
    if (this.disposed) return;
    this.dirty.add(id);
    if (this.queued) return;
    this.queued = true;
    queueMicrotask(() => {
      this.queued = false;
      this.flush();
    });
  }

  private flush(): void {
    if (this.disposed) return;
    const next = new Map(this.metrics);
    // Moving a paragraph does not change a formula's intrinsic dimensions.
    for (const id of this.dirty) {
      const entry = this.entries.get(id);
      if (!entry) {
        next.delete(id);
        continue;
      }
      // Virtual pages park roots. Their last valid measurement must remain available.
      if (
        !entry.element.isConnected ||
        entry.element.closest<HTMLElement>('[data-sciobot-block-host]')?.hidden
      ) {
        continue;
      }
      const result = this.measure(entry.element, entry.baseline, entry.zoom);
      if (result) next.set(id, Object.freeze({ ...result }));
      else next.delete(id);
    }
    this.dirty.clear();
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
      const key = `${root.dataset.sciobotInlinePage}:${id}`;
      const box = boxes.get(key),
        fallback = native.get(key);
      root.style.visibility = box && fallback ? 'visible' : 'hidden';
      root.style.overflow = 'hidden';
      if (!box || !fallback) continue;
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
    for (const observer of this.observers.values()) observer.disconnect();
    this.observers.clear();
    this.entries.clear();
    this.metrics.clear();
    this.restoreFallback();
    this.frame = null;
    setSurfaceInlineWidgetMetrics(this.surface, null);
  }
}
