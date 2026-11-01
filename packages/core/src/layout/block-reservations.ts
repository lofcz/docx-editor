// Block reservations: extra flow height a host reserves after one body paragraph, so a
// rendered widget taller than its native Word fallback pushes the following content down.
//
// A LAYOUT INPUT, never a document edit. Nothing here writes the model, enters history,
// crosses collaboration, or reaches a saved file: the surface supplies reservations to live
// passes only. Values are host measurements, so two sessions with different fonts may
// reserve differently while their documents stay identical.
//
// Units are points, the unit body flow positions in. Values are quantized to a quarter
// point so a measurement that wobbles by float noise does not relayout the document.

/** Upper bound for one reservation: a few pages. Larger answers are host bugs. */
export const MAX_BLOCK_RESERVATION_PT = 4 * 1584;
const QUANTUM = 4;

/** Upper bound for one horizontal inset: a little more than half a Letter page. */
export const MAX_BLOCK_INSET_PT = 324;

/** Layout-only room around one body paragraph or table, in points. */
export interface BlockInset {
  readonly left: number;
  readonly right: number;
  /** Raises a paragraph's `w:spacing/@before`; Word's collapse with the space after above applies. Tables ignore it. */
  readonly top?: number;
  /** Raises a paragraph's `w:spacing/@after`. Tables ignore it. */
  readonly bottom?: number;
}

/** Extra flow height after body paragraphs, keyed by canonical paragraph ID. */
export interface BlockReservations {
  /** Moves whenever any reservation or inset does; joins the section prepass memo. */
  readonly epoch: string;
  /** Points reserved after the paragraph's last line, or 0. */
  extraAfter(paragraphId: string): number;
  /** Extra indents of a body paragraph or an in-flow table, keyed by node ID. */
  insetOf?(nodeId: string): BlockInset | undefined;
}

/** Quarter-point ceiling of a finite positive value; 0 for anything else. */
export function quantizeBlockReservation(points: number): number {
  if (typeof points !== 'number' || !Number.isFinite(points) || points <= 0) return 0;
  return Math.min(MAX_BLOCK_RESERVATION_PT, Math.ceil(points * QUANTUM - 1e-6) / QUANTUM);
}

/** The cache-key fragment for one block: empty when nothing is reserved or inset. */
export function blockReservationToken(
  reservations: BlockReservations | undefined,
  nodeId: string
): string {
  const extra = reservations?.extraAfter(nodeId) ?? 0;
  const inset = reservations?.insetOf?.(nodeId);
  return (
    (extra > 0 ? `|reserve:${extra}` : '') +
    (inset ? `|inset:${inset.left},${inset.right},${inset.top ?? 0},${inset.bottom ?? 0}` : '')
  );
}

/** Narrowest text column an inset leaves: one inch. A wider inset is ignored. */
const MIN_INSET_COLUMN_PT = 72;

/** The inset when it leaves a usable column of `width` points; otherwise undefined. */
function usableInset(
  reservations: BlockReservations | undefined,
  nodeId: string,
  width: number
): BlockInset | undefined {
  const inset = reservations?.insetOf?.(nodeId);
  return inset &&
    (inset.left || inset.right) &&
    width - inset.left - inset.right >= MIN_INSET_COLUMN_PT
    ? inset
    : undefined;
}

/**
 * A body paragraph's inputs resolved against its inset column, then moved back into the full
 * column: the inset joins the indents, so every placement that reads `indent.left` and
 * `available` (lines, markers, borders, shading, boundaries) lands inside the inset.
 */
export function resolveInsetParagraph<
  T extends {
    readonly indent: { left: number; right: number; hanging: number; firstLine: number };
    readonly spacing: { readonly before: number; readonly after: number };
  },
>(
  reservations: BlockReservations | undefined,
  paragraphId: string,
  resolve: (width: number) => T,
  width: number
): T {
  const room = reservations?.insetOf?.(paragraphId);
  if (!room) return resolve(width);
  const inset = usableInset(reservations, paragraphId, width);
  const inputs = resolve(inset ? width - inset.left - inset.right : width);
  return {
    ...inputs,
    ...(inset
      ? {
          indent: {
            ...inputs.indent,
            left: inputs.indent.left + inset.left,
            right: inputs.indent.right + inset.right,
          },
        }
      : {}),
    ...(room.top || room.bottom
      ? {
          spacing: {
            ...inputs.spacing,
            before: inputs.spacing.before + (room.top ?? 0),
            after: inputs.spacing.after + (room.bottom ?? 0),
          },
        }
      : {}),
  };
}

/** The left inset a list marker moves by: the paragraph's text moved by the same room. */
export function insetLeftOf(
  reservations: BlockReservations | undefined,
  paragraphId: string,
  width: number
): number {
  return usableInset(reservations, paragraphId, width)?.left ?? 0;
}

/** The column an in-flow table paginates in, narrowed by its inset. Positioned tables pass false. */
export function insetTableColumn(
  layout: false | { readonly blockReservations?: BlockReservations | undefined },
  tableId: string,
  columnWidth: () => number,
  columnLeft: () => number
): { columnWidth: () => number; columnLeft: () => number } {
  const inset = layout ? usableInset(layout.blockReservations, tableId, columnWidth()) : undefined;
  return inset
    ? {
        columnWidth: () => columnWidth() - inset.left - inset.right,
        columnLeft: () => columnLeft() + inset.left,
      }
    : { columnWidth, columnLeft };
}

/** Quarter-point ceiling of a finite positive inset, capped; 0 for anything else. */
function quantizeInset(points: number): number {
  if (typeof points !== 'number' || !Number.isFinite(points) || points <= 0) return 0;
  return Math.min(MAX_BLOCK_INSET_PT, Math.ceil(points * QUANTUM - 1e-6) / QUANTUM);
}

/** Points reserved after one paragraph, or 0. */
export function reservedAfterOf(
  reservations: BlockReservations | undefined,
  paragraphId: string
): number {
  return reservations?.extraAfter(paragraphId) ?? 0;
}

/**
 * Mutable host-side holder. `set` and `setInsets` each replace their whole map and report
 * whether any value changed; `option` is the frozen view a layout pass reads.
 */
export function createBlockReservationState(): {
  set(next: ReadonlyMap<string, number>): boolean;
  setInsets(next: ReadonlyMap<string, BlockInset>): boolean;
  option(): BlockReservations | undefined;
} {
  let values = new Map<string, number>();
  let insets = new Map<string, BlockInset>();
  let generation = 0;
  let current: BlockReservations | undefined;
  const publish = () => {
    generation += 1;
    const frozen = values,
      frozenInsets = insets;
    current =
      frozen.size === 0 && frozenInsets.size === 0
        ? undefined
        : Object.freeze({
            epoch: `reserve:${generation}`,
            extraAfter: (paragraphId: string) => frozen.get(paragraphId) ?? 0,
            insetOf: (nodeId: string) => frozenInsets.get(nodeId),
          });
  };
  return {
    set(next) {
      const normalized = new Map<string, number>();
      for (const [id, points] of next) {
        const value = quantizeBlockReservation(points);
        if (typeof id === 'string' && id && value > 0) normalized.set(id, value);
      }
      if (
        normalized.size === values.size &&
        [...normalized].every(([id, value]) => values.get(id) === value)
      )
        return false;
      values = normalized;
      publish();
      return true;
    },
    setInsets(next) {
      const normalized = new Map<string, BlockInset>();
      for (const [id, inset] of next) {
        const left = quantizeInset(inset?.left),
          right = quantizeInset(inset?.right),
          top = quantizeInset(inset?.top ?? 0),
          bottom = quantizeInset(inset?.bottom ?? 0);
        if (typeof id === 'string' && id && (left > 0 || right > 0 || top > 0 || bottom > 0))
          normalized.set(
            id,
            Object.freeze({
              left,
              right,
              ...(top > 0 ? { top } : {}),
              ...(bottom > 0 ? { bottom } : {}),
            })
          );
      }
      if (
        normalized.size === insets.size &&
        [...normalized].every(([id, value]) => {
          const old = insets.get(id);
          return (
            old?.left === value.left &&
            old.right === value.right &&
            old.top === value.top &&
            old.bottom === value.bottom
          );
        })
      )
        return false;
      insets = normalized;
      publish();
      return true;
    },
    option: () => current,
  };
}
