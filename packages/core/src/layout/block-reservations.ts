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

/** Extra flow height after body paragraphs, keyed by canonical paragraph ID. */
export interface BlockReservations {
  /** Moves whenever any reservation does; joins the section prepass memo. */
  readonly epoch: string;
  /** Points reserved after the paragraph's last line, or 0. */
  extraAfter(paragraphId: string): number;
}

/** Quarter-point ceiling of a finite positive value; 0 for anything else. */
export function quantizeBlockReservation(points: number): number {
  if (typeof points !== 'number' || !Number.isFinite(points) || points <= 0) return 0;
  return Math.min(MAX_BLOCK_RESERVATION_PT, Math.ceil(points * QUANTUM - 1e-6) / QUANTUM);
}

/** The cache-key fragment for one paragraph: empty when nothing is reserved. */
export function blockReservationToken(
  reservations: BlockReservations | undefined,
  paragraphId: string
): string {
  const extra = reservations?.extraAfter(paragraphId) ?? 0;
  return extra > 0 ? `|reserve:${extra}` : '';
}

/** Points reserved after one paragraph, or 0. */
export function reservedAfterOf(
  reservations: BlockReservations | undefined,
  paragraphId: string
): number {
  return reservations?.extraAfter(paragraphId) ?? 0;
}

/**
 * Mutable host-side holder. `set` replaces the whole map and reports whether any value
 * changed; `option` is the frozen view a layout pass reads.
 */
export function createBlockReservationState(): {
  set(next: ReadonlyMap<string, number>): boolean;
  option(): BlockReservations | undefined;
} {
  let values = new Map<string, number>();
  let generation = 0;
  let current: BlockReservations | undefined;
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
      generation += 1;
      const frozen = values;
      current =
        frozen.size === 0
          ? undefined
          : Object.freeze({
              epoch: `reserve:${generation}`,
              extraAfter: (paragraphId: string) => frozen.get(paragraphId) ?? 0,
            });
      return true;
    },
    option: () => current,
  };
}
