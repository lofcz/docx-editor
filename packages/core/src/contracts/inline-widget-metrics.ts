// Inline widget metrics: the numbers a host renderer hands layout for one registered inline
// content control, and the rules that turn them into a box or leave Word's own metrics alone.
//
// Renderer-neutral by construction: plain numbers in, plain numbers out. Nothing here imports a
// DOM, a framework, or layout itself, so React (or any other host) measures outside layout and
// layout reads only this contract.
//
// Units are CSS pixels at 100% zoom; semantic layout converts them to points. Values use the
// layout's fixed-point scale (1/1000 px, half-to-even), so the same measurement yields the
// same box in every session that measures it, regardless of float noise in the host.

/** Measured geometry of one widget, relative to the baseline of the line it sits on. */
export interface InlineWidgetMetrics {
  /** Advance width. Greater than zero. */
  readonly width: number;
  /** Extent above the baseline. Greater than zero. */
  readonly ascent: number;
  /** Extent below the baseline. Zero or more. */
  readonly descent: number;
}

/** What layout knows about one inline control when it asks for metrics. */
export interface InlineWidgetRequest {
  /** Canonical node ID of the inline content control. */
  readonly controlId: string;
  /** The control's `w:tag`, or undefined when it has none. */
  readonly tag: string | undefined;
  /** Text Word renders for the control. Layout measures this when the widget does not apply. */
  readonly fallbackText: string;
  /** Widest box the line can hold: the content width at the control's position. */
  readonly availableWidth: number;
}

/**
 * Answers metrics for controls the host registered, or null for every other control.
 *
 * Null means "not mine, or not measured yet". Either way layout keeps the native path, so a
 * provider that has not measured a widget yet never reserves a guessed box.
 */
export type InlineWidgetMetricsProvider = (
  request: InlineWidgetRequest
) => InlineWidgetMetrics | null;

export type InlineWidgetBox =
  | {
      readonly kind: 'widget';
      readonly width: number;
      readonly ascent: number;
      readonly descent: number;
      /** True when the measured width exceeded the line and was reduced to fit it. */
      readonly clamped: boolean;
    }
  | {
      readonly kind: 'native';
      /**
       * Why the native path applies. `unregistered`: no provider, or the provider declined.
       * `invalid-metrics`: the provider answered numbers no box can have.
       */
      readonly reason: 'unregistered' | 'invalid-metrics';
    };

/** The layout fixed-point scale: one unit is 1/1000 CSS pixel. */
export const INLINE_WIDGET_METRIC_SCALE = 1000;
/** Larger answers are measurement faults, not widgets: 22 inches at 96 dpi. */
export const MAX_INLINE_WIDGET_EXTENT = 2112;

/** Round to the fixed-point grid, half to even, as layout rounds shaped advances. */
export function quantizeInlineMetric(value: number): number {
  const scaled = value * INLINE_WIDGET_METRIC_SCALE;
  const floor = Math.floor(scaled);
  const fraction = scaled - floor;
  const rounded = fraction > 0.5 || (fraction === 0.5 && floor % 2 !== 0) ? floor + 1 : floor;
  return rounded / INLINE_WIDGET_METRIC_SCALE;
}

function validExtent(value: unknown, allowZero: boolean): value is number {
  return (
    typeof value === 'number' &&
    Number.isFinite(value) &&
    (allowZero ? value >= 0 : value > 0) &&
    value <= MAX_INLINE_WIDGET_EXTENT
  );
}

/**
 * The box layout reserves for one inline control.
 *
 * Without a provider, or when it returns null, the control keeps its native Word metrics: its
 * fallback text is shaped like any other run, exactly as an unregistered control is today.
 * A provider answer that is not a valid box (non-finite, non-positive width or ascent, negative
 * descent, or beyond the extent bound) also keeps the native path rather than reserving
 * something layout cannot place. A valid box wider than the line is clamped to the line, so a
 * chip occupies at most one line and never overflows the text area.
 */
export function resolveInlineWidgetBox(
  request: InlineWidgetRequest,
  provider: InlineWidgetMetricsProvider | null | undefined
): InlineWidgetBox {
  const metrics = provider ? provider(request) : null;
  if (metrics === null || metrics === undefined) return { kind: 'native', reason: 'unregistered' };
  if (
    !validExtent(metrics.width, false) ||
    !validExtent(metrics.ascent, false) ||
    !validExtent(metrics.descent, true)
  )
    return { kind: 'native', reason: 'invalid-metrics' };
  const width = quantizeInlineMetric(metrics.width);
  const available =
    Number.isFinite(request.availableWidth) && request.availableWidth > 0
      ? quantizeInlineMetric(request.availableWidth)
      : width;
  return {
    kind: 'widget',
    width: Math.min(width, available),
    ascent: quantizeInlineMetric(metrics.ascent),
    descent: quantizeInlineMetric(metrics.descent),
    clamped: width > available,
  };
}
