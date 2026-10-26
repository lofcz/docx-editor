import { TreePackageStore, type OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { EDIT_PLAN_LIMITS } from '../diff/plan';
import { readReviewGroups } from './groups';
import {
  planOwnedProposalRemoval,
  type OwnedProposalGroup,
  type OwnedProposalRemovalResult,
} from './owned-proposal-plan';
import { ownedProposalSelectionProblem } from './owned-proposal-input';

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/**
 * Select currently pending groups from an explicit host-retained ownership ledger.
 * Missing groups stay in the caller's ledger so native undo can make them pending again.
 * R24 validates inverses inside an unobserved native frame, isolating its preview primitives.
 * Restored group visibility and matching author names never grant ownership.
 */
export function reconcileRecompositionOwnership(
  pkg: OoxmlPackage,
  revision: RevisionToken,
  ledger: readonly OwnedProposalGroup[]
):
  | { readonly success: true; readonly selection: readonly OwnedProposalGroup[] }
  | NativeEditFailure {
  try {
    const copied = copyBoundedJson({ revision, ledger }, EDIT_PLAN_LIMITS);
    if (!copied.ok) return fail(copied.reason);
    const input = copied.value as unknown as {
      revision: RevisionToken;
      ledger: readonly OwnedProposalGroup[];
    };
    const problem = ownedProposalSelectionProblem(input.ledger);
    if (problem) return fail(problem);
    const read = readReviewGroups(pkg);
    if (!read.success) return read;
    const active = new Set(read.groups.map((group) => group.id));
    const selection = input.ledger.filter((group) => active.has(group.groupId));
    const part = pkg.parts.get(pkg.mainDocumentPart);
    if (!part) return fail('The native document body is missing.', 'invalid-document');
    const isolated = new TreePackageStore(pkg, part);
    let planned: OwnedProposalRemovalResult | undefined;
    const checked = isolated.transact(
      { kind: 'body' },
      () => {
        planned = planOwnedProposalRemoval({ pkg, revision: input.revision }, selection);
      },
      { recordsHistory: false }
    );
    if (!checked.ok || !planned)
      return fail('Unable to validate owned proposal inverses.', 'invalid-document');
    if (!planned.success) return planned;
    // R24 returns detached, deeply frozen receipts after exact current membership validation.
    return Object.freeze({ success: true, selection: planned.plan.selection });
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'invalid-document');
  }
}

export function freezeRecompositionOwnership(group: OwnedProposalGroup): OwnedProposalGroup {
  return Object.freeze({
    groupId: group.groupId,
    native: Object.freeze(
      group.native.map((member) =>
        Object.freeze({
          partName: member.partName,
          addresses: Object.freeze(
            member.addresses.map((address) => Object.freeze({ ...address }))
          ),
        })
      )
    ),
    properties: Object.freeze(group.properties.map((member) => Object.freeze({ ...member }))),
  });
}
