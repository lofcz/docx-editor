import type { ContentLockPolicy } from '../package/ooxml-edit.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import type { RevisionIdReservationSource } from './revision-id-reservations.ts';

export type NoteCascadeFn = (before: OoxmlPackage, after: OoxmlPackage) => OoxmlPackage | null;

/** How a package store is constructed: limits, history depth, and review contributions. */
export interface TreePackageStoreOptions {
  /** Pure host metadata read before each transaction. Retained IDs cannot be newly allocated. */
  readonly reservedRevisionIds?: RevisionIdReservationSource;
  readonly historyLimit?: number;
  /** Opt-in relaxation of `w:lock` for run-property writes, for the body and every story store. */
  readonly contentLockPolicy?: ContentLockPolicy;
  /** Bound on opened story stores; defaults to {@link DEFAULT_MAX_EDITABLE_STORY_PARTS}. */
  readonly maxEditableStoryParts?: number;
  /**
   * Test seam for note-reference cascade after `deleteText` / `deleteBlock`. Production uses
   * {@link cascadeDeletedNoteReferences}.
   */
  readonly cascadeDeletedNoteReferences?: NoteCascadeFn;
}
