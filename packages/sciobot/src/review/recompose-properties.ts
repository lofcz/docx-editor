import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { TreePackageStore } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { planDomainProperties, type DomainPropertyTarget } from '../diff/domain-properties';
import { createEditPlan, EDIT_PLAN_LIMITS } from '../diff/plan';
import { jsonDataKey } from '../diff/target-key';
import { reviewGroupRevisionReservations } from './group-allocation';
import {
  verifyOwnedProposalContentProofs,
  type OwnedProposalContentProof,
} from './owned-proposal-content';
import { planOwnedProposalRemoval, type OwnedProposalGroup } from './owned-proposal-plan';
import { stageOwnedProposalRemoval } from './owned-proposal-stage';
import {
  reconcileRecompositionOwnership,
  freezeRecompositionOwnership,
} from './recompose-ownership';
import {
  replaceOwnedProposals,
  type OwnedProposalReplacementResult,
} from './replace-owned-proposals';
import type { PropertyFallbackTarget } from './track-properties';
import type { TrackTextContext } from './tracked-inline';
import { readTrackedStructureInput } from './tracked-structure';

export interface RecomposePropertiesRequest {
  readonly revision: RevisionToken;
  readonly blockId: string;
  /** Complete desired values for the fields this gesture owns. Omitted fields remain unchanged. */
  readonly target: DomainPropertyTarget;
  readonly ownership: readonly OwnedProposalGroup[];
  readonly contentProofs: readonly OwnedProposalContentProof[];
  readonly proposalId: string;
  /** Exact generated label before owned proposals, and its desired localized replacement. */
  readonly fallback: PropertyFallbackTarget;
}

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'invalid-input',
  error,
});

/**
 * Replace explicitly owned manual field proposals, without parsing or rebuilding native content.
 * AI target planning retains its separate domain identity rules. This synchronous gesture never
 * grants ownership from visible groups, authors, or serialized block identifiers alone.
 */
export function recomposeOwnedProperties(
  editor: DocxEditorInstance,
  request: RecomposePropertiesRequest,
  context: TrackTextContext
): OwnedProposalReplacementResult {
  try {
    const copied = copyBoundedJson({ request, context }, EDIT_PLAN_LIMITS);
    if (!copied.ok) return fail(copied.reason);
    const input = copied.value as unknown as {
      request: RecomposePropertiesRequest;
      context: TrackTextContext;
    };
    const data = input.request;
    if (
      !data ||
      !Array.isArray(data.ownership) ||
      !Array.isArray(data.contentProofs) ||
      data.fallback?.blockId !== data.blockId ||
      Object.keys(data).some(
        (key) =>
          ![
            'revision',
            'blockId',
            'target',
            'ownership',
            'contentProofs',
            'proposalId',
            'fallback',
          ].includes(key)
      )
    )
      return fail('Invalid owned field proposal request.');
    const token = createEditPlan(data.revision, []);
    if (!token.success) return token;
    const checked = readTrackedStructureInput(editor, token.plan, input.context, 'field proposal');
    if (!checked.success) return checked;
    const { surface, actual, allowGroupInitialization } = checked.input;
    const pkg = surface.session.currentPackage();
    const selected = reconcileRecompositionOwnership(pkg, actual, data.ownership);
    if (!selected.success) return selected;
    const ids = new Set(data.ownership.map((group) => group.groupId));
    if (data.contentProofs.some((proof) => !ids.has(proof.groupId)))
      return fail('Content evidence names a group outside the ownership ledger.');
    const proofRefusal = verifyOwnedProposalContentProofs(
      pkg,
      actual,
      selected.selection,
      data.contentProofs
    );
    if (proofRefusal) return proofRefusal;
    // A field control must not gain authority over unrelated native body proposals.
    if (
      selected.selection.some(
        (group) => group.properties.length !== 1 || group.properties[0]!.blockId !== data.blockId
      )
    )
      return fail('Field proposal ownership must address this one application block.');
    const removal = planOwnedProposalRemoval({ pkg, revision: actual }, selected.selection);
    if (!removal.success) return removal;
    const store = new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!, {
      reservedRevisionIds: reviewGroupRevisionReservations,
    });
    const staged = store.transact(
      { kind: 'body' },
      (ctx) =>
        stageOwnedProposalRemoval(ctx, removal.plan, allowGroupInitialization, (run) => {
          if (!run()) throw Error('A detached field proposal inverse refused.');
        }),
      { recordsHistory: false, actorId: surface.collaborationSession()?.identity.actorId }
    );
    if (!staged.ok) return fail(`Unable to prepare the field proposal baseline: ${staged.reason}.`);
    const planned = planDomainProperties(
      { pkg: store.currentPackage(), revision: actual, blockId: data.blockId },
      data.target,
      'field-edit'
    );
    if (!planned.success) return planned;
    const plan = createEditPlan(
      actual,
      planned.plan.operations.filter(
        (edit) =>
          edit.kind !== 'payload' ||
          edit.before.present !== edit.after.present ||
          (edit.before.present &&
            edit.after.present &&
            jsonDataKey(edit.before.value) !== jsonDataKey(edit.after.value))
      )
    );
    if (!plan.success) return plan;
    const result = replaceOwnedProposals(
      editor,
      selected.selection,
      plan.plan,
      input.context,
      plan.plan.operations.length
        ? {
            properties: {
              proposalId: data.proposalId,
              intent: 'field-edit',
              fallbackTargets: [data.fallback],
            },
          }
        : {}
    );
    if (!result.success) return result;
    const ledger = new Map(data.ownership.map((group) => [group.groupId, group]));
    for (const group of result.ownership) ledger.set(group.groupId, group);
    const proofs = new Map(data.contentProofs.map((proof) => [proof.groupId, proof]));
    for (const proof of result.contentProofs) proofs.set(proof.groupId, proof);
    return {
      ...result,
      ownership: Object.freeze([...ledger.values()].map(freezeRecompositionOwnership)),
      contentProofs: Object.freeze(
        [...proofs.values()].map((proof) => Object.freeze({ ...proof }))
      ),
    };
  } catch (cause) {
    return {
      success: false,
      code: 'apply-failed',
      error: cause instanceof Error ? cause.message : String(cause),
    };
  }
}
