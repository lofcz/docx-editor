import type { TextMeasurer } from '../layout/text-measurer.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';
import { refreshCompositionActive } from './refresh-write-guard.ts';

/** One immutable host measurement revision. @public */
export type InlineWidgetMetricsSnapshot = NonNullable<TextMeasurer['inlineWidgets']>;

function validateSnapshot(next: InlineWidgetMetricsSnapshot | null): void {
  if (next !== null && (typeof next?.measure !== 'function' || typeof next.revision !== 'string'))
    throw new TypeError('Inline widget measurements need a provider and revision string.');
}

/** Local measurement decoration survives font replacement without changing native measurer behavior. */
export class SurfaceInlineWidgetMetrics {
  private snapshot: InlineWidgetMetricsSnapshot | null = null;
  private readonly wrappers = new WeakMap<
    TextMeasurer,
    { snapshot: InlineWidgetMetricsSnapshot; measurer: TextMeasurer }
  >();
  private readonly originals = new WeakMap<TextMeasurer, TextMeasurer>();

  set(next: InlineWidgetMetricsSnapshot | null): boolean {
    validateSnapshot(next);
    if (
      next === this.snapshot ||
      (next &&
        this.snapshot &&
        next.measure === this.snapshot.measure &&
        next.revision === this.snapshot.revision)
    )
      return false;
    this.snapshot = next ? Object.freeze({ measure: next.measure, revision: next.revision }) : null;
    return true;
  }

  native(measurer: TextMeasurer): TextMeasurer {
    return this.originals.get(measurer) ?? measurer;
  }

  measurer(base: TextMeasurer): TextMeasurer {
    base = this.native(base);
    const snapshot = this.snapshot;
    if (!snapshot) return base;
    const cached = this.wrappers.get(base);
    if (cached?.snapshot === snapshot) return cached.measurer;
    const measured: TextMeasurer = {
      measure: (text, style) => base.measure(text, style),
      lineMetrics: (style, text) => base.lineMetrics(style, text),
      ...(base.caretAdvances
        ? { caretAdvances: (text, style) => base.caretAdvances!(text, style) }
        : {}),
      ...(base.inkBounds ? { inkBounds: (text, style) => base.inkBounds!(text, style) } : {}),
      ...(base.textOutline ? { textOutline: (text, style) => base.textOutline!(text, style) } : {}),
      ...(base.hasResolvedFont ? { hasResolvedFont: (style) => base.hasResolvedFont!(style) } : {}),
      ...(base.strikeoutMetrics
        ? { strikeoutMetrics: (style) => base.strikeoutMetrics!(style) }
        : {}),
      inlineWidgets: snapshot,
    };
    this.wrappers.set(base, { snapshot, measurer: measured });
    this.originals.set(measured, base);
    return measured;
  }
}

const updates = new WeakMap<
  PaginatedSurface,
  (snapshot: InlineWidgetMetricsSnapshot | null) => void
>();
interface PendingMetrics {
  snapshot: InlineWidgetMetricsSnapshot | null;
  readonly target: HTMLElement;
  readonly complete: () => void;
}
const pending = new WeakMap<PaginatedSurface, PendingMetrics>();

/** Register the native surface's local measurement lifetime. @internal */
export function registerSurfaceInlineWidgetMetrics(
  surface: PaginatedSurface,
  update: (snapshot: InlineWidgetMetricsSnapshot | null) => void
): void {
  updates.set(surface, update);
}

/**
 * Publish or clear host-only inline geometry. This creates no native edit or undo item.
 * Keep each snapshot and provider result stable until the next publication.
 * During native IME composition, only the latest snapshot applies after readback completes.
 * @public
 */
export function setSurfaceInlineWidgetMetrics(
  surface: PaginatedSurface,
  snapshot: InlineWidgetMetricsSnapshot | null
): void {
  const update = updates.get(surface);
  if (!update) throw new Error('The native surface does not support inline widget measurements.');
  validateSnapshot(snapshot);
  if (refreshCompositionActive(surface)) {
    const frozen = snapshot
      ? Object.freeze({ measure: snapshot.measure, revision: snapshot.revision })
      : null;
    const previous = pending.get(surface);
    if (previous) {
      previous.snapshot = frozen;
      return;
    }
    const target = surface.hostElement().querySelector<HTMLElement>('.docx-pages');
    if (!target) throw new Error('The native composition surface is missing.');
    const entry: PendingMetrics = {
      snapshot: frozen,
      target,
      complete: () =>
        queueMicrotask(() => {
          if (pending.get(surface) !== entry) return;
          pending.delete(surface);
          target.removeEventListener('compositionend', entry.complete, true);
          setSurfaceInlineWidgetMetrics(surface, entry.snapshot);
        }),
    };
    pending.set(surface, entry);
    target.addEventListener('compositionend', entry.complete, { capture: true, once: true });
    return;
  }
  const previous = pending.get(surface);
  if (previous) {
    pending.delete(surface);
    previous.target.removeEventListener('compositionend', previous.complete, true);
  }
  surface.flushPendingInput();
  update(snapshot);
}
