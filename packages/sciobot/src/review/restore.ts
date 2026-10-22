import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { DEFAULT_PAYLOAD_LIMITS } from '../blocks/payload-schema';
import { copyBoundedJson } from '../data/json';
import { readReviewGroups, type NativeReviewGroup } from './groups';
import type { PendingPropertyProposal } from './property-projection';
import { readStoredPropertyProposals } from './stored-property-proposals';

/** A read projection of the loaded DOCX. It stores no chat or document session state. */
export interface RestoredReviewState {
  readonly revision: RevisionToken;
  readonly groups: readonly NativeReviewGroup[];
  readonly pendingProperties: readonly PendingPropertyProposal[];
  /** Pending read-only groups remain visible until their ownership is reconciled. */
  readonly hasPendingReview: boolean;
  /** Intrinsic group eligibility. Native permissions and history gates still apply on decisions. */
  readonly reviewableGroupIds: readonly string[];
}

export type RestoreReviewResult =
  | { readonly success: true; readonly state: RestoredReviewState }
  | NativeEditFailure;

/**
 * Recover review availability from persisted native/application records without writes.
 * Supply the identity and revision of the current load. Saved canonical node IDs are never reused.
 */
export function restoreReviewState(
  pkg: OoxmlPackage,
  revision: RevisionToken
): RestoreReviewResult {
  const copied = copyBoundedJson(revision, DEFAULT_PAYLOAD_LIMITS);
  if (!copied.ok) return { success: false, code: 'invalid-input', error: copied.reason };
  const token = copied.value as unknown as RevisionToken;
  if (
    !token ||
    typeof token.documentId !== 'string' ||
    !token.documentId.trim() ||
    typeof token.epoch !== 'string' ||
    !token.epoch.trim() ||
    !Number.isSafeInteger(token.revision) ||
    token.revision < 0
  )
    return { success: false, code: 'invalid-input', error: 'Invalid loaded document revision.' };
  const groups = readReviewGroups(pkg);
  if (!groups.success) return groups;
  const current = Object.freeze({
    documentId: token.documentId,
    epoch: token.epoch,
    revision: token.revision,
  });
  const properties = readStoredPropertyProposals(pkg, current);
  if (!properties.success) return properties;
  return Object.freeze({
    success: true,
    state: Object.freeze({
      revision: current,
      groups: groups.groups,
      pendingProperties: properties.pending,
      hasPendingReview: groups.groups.length > 0,
      reviewableGroupIds: Object.freeze(
        groups.groups.filter((group) => !group.readOnly).map((group) => group.id)
      ),
    }),
  });
}
