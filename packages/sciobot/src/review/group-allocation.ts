import { decodeBlockPresenceRecord } from './block-presence-schema';
import { readInlinePresenceExtension } from './inline-presence-native';
import {
  runWithRevisionIdReservations,
  type OoxmlPackage,
  type RevisionIdReservation,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { readReviewGroupStore } from './group-record-store';
import { readPropertyReviewStore } from './property-record-store';

/** Native addresses remain reserved after decisions and record removal, for either actor's undo. */
export function reviewGroupRevisionReservations(
  pkg: OoxmlPackage
): readonly RevisionIdReservation[] {
  const store = readReviewGroupStore(pkg);
  if (store.status === 'invalid') throw new Error(store.reason);
  const ids = new Map<string, Set<string>>();
  for (const entry of store.status === 'ready' ? store.retainedRecords.values() : []) {
    if (entry.decoded.status === 'payload-decision') continue;
    if (entry.decoded.status !== 'supported')
      throw new Error('Unsupported persisted review group version.');
    for (const member of entry.decoded.value.members) {
      const part = ids.get(member.partName) ?? new Set<string>();
      ids.set(member.partName, part);
      part.add(member.revision.id);
    }
  }
  const properties = readPropertyReviewStore(pkg);
  if (properties.status === 'invalid') throw new Error(properties.reason);
  if (properties.status === 'ready')
    for (const entry of properties.retainedRecords.values()) {
      if (entry.decoded.status !== 'supported') {
        const presence = decodeBlockPresenceRecord(entry.raw);
        if (presence.status === 'supported') {
          for (const owner of readInlinePresenceExtension(presence.value)?.values() ?? []) {
            const part = ids.get(owner.partName) ?? new Set<string>();
            ids.set(owner.partName, part);
            for (const address of owner.gap ?? [owner.revision]) part.add(address.id);
          }
          continue;
        }
        throw new Error('Unsupported property proposal version.');
      }
      for (const address of entry.decoded.value.nativeRevisions) {
        const part = ids.get(address.partName) ?? new Set<string>();
        ids.set(address.partName, part);
        part.add(address.id);
      }
    }
  return Object.freeze(
    [...ids].map(([partName, values]) =>
      Object.freeze({ partName, ids: Object.freeze([...values]) })
    )
  );
}

/** Use the same reservations for pure executor preflights as the review module uses for commits. */
export function withReviewGroupRevisionReservations<T>(
  pkg: OoxmlPackage | null,
  run: () => T
): T | NativeEditFailure {
  if (!pkg) return run(); // The executor retains its document-not-ready validation.
  let reservations: readonly RevisionIdReservation[];
  try {
    reservations = reviewGroupRevisionReservations(pkg);
  } catch (cause) {
    return {
      success: false,
      code: 'invalid-document',
      error: cause instanceof Error ? cause.message : String(cause),
    };
  }
  return runWithRevisionIdReservations(reservations, run);
}
