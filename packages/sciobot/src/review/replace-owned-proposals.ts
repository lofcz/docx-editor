import { nativeNumberingInfrastructureReady } from '../native/numbering-infrastructure';
import { stageTrackedPresence } from './tracked-stage-presence';
import { stageTrackedPayloads, type TrackedPayloadOptions } from './tracked-stage-payloads';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  TreePackageStore,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  type OoxmlPackage,
  type TreeTransactionContext,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { decodeEditPlan, EDIT_PLAN_LIMITS, type CanonicalEditPlan } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import { withReviewDecisionGuard } from './decision-guard';
import {
  reviewGroupRevisionReservations,
  withReviewGroupRevisionReservations,
} from './group-allocation';
import { captureReviewGroupHistory, reviewGroupHistoryRefusal } from './group-history';
import { groupRemovalBudget } from './group-removal-budget';
import { groupReviewFieldBudget } from './group-field-budget';
import { persistReviewGroups, readReviewGroups } from './groups';
import { planOwnedProposalRemoval, type OwnedProposalGroup } from './owned-proposal-plan';
import { stageOwnedProposalRemoval } from './owned-proposal-stage';
import {
  captureOwnedProposalContentProofs,
  type OwnedProposalContentProof,
} from './owned-proposal-content';
import { canonicalTrackedDate, type TrackTextContext } from './tracked-inline';
import { NativeWriteRefusal } from '../diff/fragment-seeds';
import { stageTrackedPlan } from './tracked-stage';
import { currentRevisionGuardContent } from './revision-property-content';

export type OwnedProposalReplacementResult =
  | {
      readonly success: true;
      readonly committed: boolean;
      readonly revision: RevisionToken;
      /** Retain these exact receipts as ownership for the next replacement. */
      readonly ownership: readonly OwnedProposalGroup[];
      /** Semantic evidence captured before publication, for safe later-wave replacement. */
      readonly contentProofs: readonly OwnedProposalContentProof[];
    }
  | NativeEditFailure;

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
class ReplacementRefusal extends Error {
  constructor(readonly failure: NativeEditFailure) {
    super(failure.error);
  }
}
const requireSuccess = <T extends { readonly success: boolean }>(
  value: T
): Exclude<T, NativeEditFailure> => {
  if (!value.success) throw new ReplacementRefusal(value as unknown as NativeEditFailure);
  return value as Exclude<T, NativeEditFailure>;
};

/** Remaining revisions belong to other proposals, even when their authors match. */
function preserveRemainingRevisions(before: OoxmlPackage, after: OoxmlPackage): void {
  for (const [name, part] of before.parts) {
    const remaining = collectRevisionSites(part);
    if (!remaining.length) continue;
    const next = after.parts.get(name);
    const sites = new Map(
      next ? collectRevisionSites(next).map((site) => [site.node.id, site.node]) : []
    );
    for (const site of remaining) {
      const node = sites.get(site.node.id);
      const priorContent = currentRevisionGuardContent(part, site.node);
      const nextContent = next && node ? currentRevisionGuardContent(next, node) : null;
      if (
        !node ||
        canonicalOoxmlFingerprint(node) !== canonicalOoxmlFingerprint(site.node) ||
        (priorContent !== null &&
          (!nextContent ||
            canonicalOoxmlFingerprint(priorContent) !== canonicalOoxmlFingerprint(nextContent)))
      )
        throw new ReplacementRefusal(
          fail('Replacement would change an unselected revision.', 'ambiguous-target')
        );
    }
  }
}

function preserveRemainingGroups(before: OoxmlPackage, after: OoxmlPackage): void {
  const original = requireSuccess(readReviewGroups(before)).groups;
  const current = new Map(
    requireSuccess(readReviewGroups(after)).groups.map((group) => [group.id, group])
  );
  const members = (group: (typeof original)[number]) =>
    JSON.stringify([
      group.native
        .map((member) =>
          JSON.stringify([member.partName, member.address, [...member.siteNodeIds].sort()])
        )
        .sort(),
      group.properties.map((member) => JSON.stringify([member.proposalId, member.blockId])).sort(),
    ]);
  for (const group of original) {
    const next = current.get(group.id);
    if (!next || group.readOnly !== next.readOnly || members(group) !== members(next))
      throw new ReplacementRefusal(
        fail('Replacement would change an unselected review group.', 'ambiguous-target')
      );
  }
}

/**
 * Replace exact host-owned pending groups and propose ordered edits in one native transaction.
 * The replacement plan uses coordinates after the selected inverses, with the original revision.
 * The caller prepares it from a detached read; this command validates every old value again.
 * Every dependent edit sees the current staged package. No document snapshot is installed.
 */
export function replaceOwnedProposals(
  editor: DocxEditorInstance,
  selection: readonly OwnedProposalGroup[],
  replacement: CanonicalEditPlan,
  context: TrackTextContext,
  payloads: TrackedPayloadOptions = {}
): OwnedProposalReplacementResult {
  return withReviewDecisionGuard(editor, () =>
    commit(editor, selection, replacement, context, payloads)
  );
}

function commit(
  editor: DocxEditorInstance,
  selection: readonly OwnedProposalGroup[],
  replacement: CanonicalEditPlan,
  context: TrackTextContext,
  payloads: TrackedPayloadOptions
): OwnedProposalReplacementResult {
  const copied = copyBoundedJson({ selection, replacement, context, payloads }, EDIT_PLAN_LIMITS);
  if (!copied.ok) return fail(copied.reason);
  const input = copied.value as unknown as {
    selection: readonly OwnedProposalGroup[];
    replacement: CanonicalEditPlan;
    context: TrackTextContext;
    payloads: TrackedPayloadOptions;
  };
  const attribution = input.context;
  if (
    !attribution ||
    typeof attribution.documentId !== 'string' ||
    !attribution.documentId ||
    typeof attribution.epoch !== 'string' ||
    !attribution.epoch ||
    typeof attribution.author !== 'string' ||
    !attribution.author.trim() ||
    attribution.author.length > 4096 ||
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(attribution.author) ||
    /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(
      attribution.author
    ) ||
    typeof attribution.date !== 'string'
  )
    return fail('Invalid owned-proposal replacement context.');
  const date = canonicalTrackedDate(attribution.date);
  if (!date) return fail('Replacement proposals need a valid UTC ISO-8601 date.');
  const decoded = decodeEditPlan(JSON.stringify(input.replacement));
  if (!decoded.success) return decoded;
  const plan = decoded.plan;
  const surface = editor.surface;
  if (!surface) return fail('The native editor is not attached.', 'document-not-ready');
  if (surface.editingMode() === 'view')
    return fail('The document is open for viewing.', 'unsupported-operation');
  const historyRefusal = reviewGroupHistoryRefusal(surface.collaborationSession());
  if (historyRefusal) return fail(historyRefusal, 'unsupported-operation');
  const capability = editor.can({ type: 'setReviewDisplayMode', mode: 'proposed' });
  if (!capability.ok) return fail(String(capability.reason), 'unsupported-operation');
  let result: OwnedProposalReplacementResult | undefined;
  try {
    const refused = surface.commitReviewOps((historyToken) => {
      const current = readEditorRevision(editor, attribution.documentId, attribution.epoch);
      if (!current || editor.surface !== surface) {
        result = fail('The native editor is not attached.', 'document-not-ready');
        return { committed: false, reason: result.error };
      }
      if (!sameRevision(plan.revision, current)) {
        result = {
          ...fail('Read the current document before replacing owned proposals.', 'stale-revision'),
          expectedRevision: plan.revision,
          actualRevision: current,
        };
        return { committed: false, reason: result.error };
      }
      const before = surface.session.currentPackage();
      if (
        editor.snapshot().collaborationStatus !== 'inactive' &&
        !nativeNumberingInfrastructureReady(before) &&
        plan.operations.some(
          (edit) =>
            edit.kind === 'siblings' &&
            !!edit.blocks?.length &&
            edit.lists?.some(
              (binding) => binding.target.kind === 'ordered' || binding.target.kind === 'bullet'
            )
        )
      )
        throw new ReplacementRefusal(
          fail(
            'Prepare native numbering infrastructure before sharing application list fragments.',
            'unsupported-operation'
          )
        );

      // The planner uses native operations on detached parts. Run it before the observed
      // transaction begins, so its inverse previews cannot enter the live operation journal.
      const removal = requireSuccess(
        planOwnedProposalRemoval({ pkg: before, revision: current }, input.selection)
      ).plan;
      const allowInitialization = editor.snapshot().collaborationStatus === 'inactive';
      let ownership: readonly OwnedProposalGroup[] = [];
      let contentProofs: readonly OwnedProposalContentProof[] = [];
      let flags = 0,
        fields = 0;
      const stage = (ctx: TreeTransactionContext, measure: boolean) => {
        const original = ctx.currentPackage();
        if (!measure && original !== before)
          throw new ReplacementRefusal(
            fail('The document changed before replacement staging.', 'stale-revision')
          );
        const step = (run: () => boolean) => {
          const prior = ctx.currentPackage();
          if (!run()) throw new Error('A native owned-proposal operation refused.');
          if (measure) {
            flags += groupRemovalBudget(prior, ctx.currentPackage());
            fields += groupReviewFieldBudget(prior, ctx.currentPackage());
          }
        };
        const edit = (apply: (pkg: OoxmlPackage) => OoxmlPackage) =>
          step(() => ctx.applyPackage(apply));
        stageOwnedProposalRemoval(ctx, removal, allowInitialization, step);
        const middle = ctx.currentPackage();
        const reserved = withReviewGroupRevisionReservations(middle, () => {
          const receipts = stageTrackedPlan(
            {
              currentPackage: () => ctx.currentPackage(),
              apply: (partName, op) => step(() => ctx.applyTo(partName, op)),
              edit,
            },
            { ...plan, operations: plan.operations.filter((edit) => edit.kind !== 'payload') },
            { author: attribution.author, date },
            !surface.session.trackingSettings().doNotTrackFormatting &&
              editor.snapshot().revisionMarkup.trackFormatting
          );
          stageTrackedPayloads(
            {
              currentPackage: () => ctx.currentPackage(),
              apply: (partName, op) => step(() => ctx.applyTo(partName, op)),
              edit,
            },
            plan,
            { author: attribution.author, date },
            input.payloads
          );
          stageTrackedPresence(
            {
              currentPackage: () => ctx.currentPackage(),
              apply: (partName, op) => step(() => ctx.applyTo(partName, op)),
              edit,
            },
            middle,
            plan,
            { author: attribution.author, date },
            input.payloads.presence
          );
          preserveRemainingRevisions(middle, ctx.currentPackage());
          edit((working) => persistReviewGroups(middle, working, receipts, allowInitialization));
          preserveRemainingGroups(middle, ctx.currentPackage());
          return { success: true as const };
        });
        requireSuccess(reserved);
        const prior = new Set(
          requireSuccess(readReviewGroups(middle)).groups.map((group) => group.id)
        );
        ownership = Object.freeze(
          requireSuccess(readReviewGroups(ctx.currentPackage()))
            .groups.filter((group) => !prior.has(group.id))
            .map(
              (group): OwnedProposalGroup =>
                Object.freeze({
                  groupId: group.id,
                  native: Object.freeze(
                    group.native.map((member) =>
                      Object.freeze({
                        partName: member.partName,
                        addresses: Object.freeze([Object.freeze({ ...member.address })]),
                      })
                    )
                  ),
                  properties: Object.freeze(
                    group.properties.map((member) => Object.freeze({ ...member }))
                  ),
                })
            )
        );
        contentProofs = requireSuccess(
          captureOwnedProposalContentProofs(ctx.currentPackage(), current, ownership)
        ).contentProofs;
      };
      // A detached native coordinator also previews cleanup and package cascades. It has no live
      // subscriber, collaboration binding, or history. Only ordinary operations enter the live store.
      const preview = new TreePackageStore(before, before.parts.get(before.mainDocumentPart)!, {
        reservedRevisionIds: reviewGroupRevisionReservations,
      });
      let last = before;
      const checked = preview.transact(
        { kind: 'body' },
        (ctx) => {
          stage(ctx, true);
          last = ctx.currentPackage();
        },
        { recordsHistory: false }
      );
      if (!checked.ok) throw new Error(`Replacement preflight refused: ${checked.reason}.`);
      preserveRemainingRevisions(last, preview.currentPackage());
      flags += groupRemovalBudget(last, preview.currentPackage());
      fields += groupReviewFieldBudget(last, preview.currentPackage());
      ownership = [];
      contentProofs = [];
      const applied = captureReviewGroupHistory(
        surface.collaborationSession(),
        () =>
          surface.session.applyTreeOps(
            [],
            undefined,
            undefined,
            { kind: 'body' },
            {
              historyToken,
              stage: (ctx) => stage(ctx, false),
            }
          ),
        flags,
        fields
      );
      if (applied.rejected) {
        result = fail(
          String(applied.reason ?? 'Owned-proposal replacement refused.'),
          'apply-failed'
        );
      } else {
        result = {
          success: true,
          committed: applied.committed,
          revision: Object.freeze({ ...current, revision: surface.session.packageRevision() }),
          ownership,
          contentProofs,
        };
      }
      return applied;
    }, 'revision-resolve');
    return (
      result ??
      fail(String(refused ?? 'Owned-proposal replacement did not commit.'), 'unsupported-operation')
    );
  } catch (cause) {
    return cause instanceof ReplacementRefusal
      ? cause.failure
      : cause instanceof NativeWriteRefusal
        ? fail(cause.message, cause.code)
        : fail(cause instanceof Error ? cause.message : String(cause), 'apply-failed');
  }
}
