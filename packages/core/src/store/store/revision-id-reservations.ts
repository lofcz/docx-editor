import { MAX_DECIMAL_ID } from '../package/actor-scoped-ids.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';

/** Native revision IDs retained outside live Word revision elements. @public */
export interface RevisionIdReservation {
  readonly partName: string;
  readonly ids: readonly string[];
}

/** Pure host read of retained IDs. Core does not interpret the host's metadata. @public */
export type RevisionIdReservationSource = (pkg: OoxmlPackage) => readonly RevisionIdReservation[];

/** Resource bounds for one synchronous allocation scope. @public */
export const REVISION_ID_RESERVATION_LIMITS = Object.freeze({
  maxParts: 10_000,
  maxIds: 250_000,
  maxStringLength: 4096,
});

let active: ReadonlyMap<string, ReadonlySet<string>> | undefined;
const EMPTY: ReadonlySet<string> = new Set();

/** Normalize only values that can equal a newly allocated nonnegative Word decimal ID. */
function decimal(value: string): string | null {
  const digits = /^\+?0*(\d{1,10})$/.exec(value);
  if (!digits) return null;
  const number = Number(digits[1]);
  return number <= MAX_DECIMAL_ID ? String(number) : null;
}

/** Internal read shared by native revision allocation and clipboard revision allocation. */
export function reservedRevisionIdsOf(partName: string): ReadonlySet<string> {
  return active?.get(partName) ?? EMPTY;
}

/**
 * Run synchronous native allocation with an owned copy of retained IDs. Live IDs remain occupied.
 * Nondecimal or out-of-range strings cannot equal newly minted IDs and are ignored.
 * Nested scopes replace their caller's reservations and restore them on return or throw.
 * Do not return asynchronous work: reservations end when this callback returns.
 * This host infrastructure scope writes no document, history, or collaboration state.
 * @public
 */
export function runWithRevisionIdReservations<T>(
  reservations: readonly RevisionIdReservation[],
  run: () => T
): T {
  if (!Array.isArray(reservations) || reservations.length > REVISION_ID_RESERVATION_LIMITS.maxParts)
    throw new TypeError('Invalid revision ID reservation parts.');
  let count = 0;
  const owned = new Map<string, Set<string>>();
  for (const entry of reservations) {
    if (
      !entry ||
      typeof entry.partName !== 'string' ||
      !entry.partName.startsWith('/') ||
      entry.partName.length > REVISION_ID_RESERVATION_LIMITS.maxStringLength ||
      /[\\\u0000-\u001f]/.test(entry.partName) ||
      entry.partName.split('/').some((segment: string) => segment === '..' || segment === '.') ||
      !Array.isArray(entry.ids)
    )
      throw new TypeError('Invalid revision ID reservation entry.');
    count += entry.ids.length;
    if (count > REVISION_ID_RESERVATION_LIMITS.maxIds)
      throw new TypeError('Revision ID reservations exceed their limit.');
    const ids = owned.get(entry.partName) ?? new Set<string>();
    owned.set(entry.partName, ids);
    for (const raw of entry.ids) {
      if (typeof raw !== 'string' || raw.length > REVISION_ID_RESERVATION_LIMITS.maxStringLength)
        throw new TypeError('Invalid reserved revision ID.');
      const id = decimal(raw);
      if (id !== null) ids.add(id);
    }
  }
  const previous = active;
  active = owned;
  try {
    return run();
  } finally {
    active = previous;
  }
}
