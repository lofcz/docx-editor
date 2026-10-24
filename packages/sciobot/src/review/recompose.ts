import { isCurrentNativeTarget } from './recompose-current-target';
import {
  prepareInlineImageResources,
  type InlineImageResolver,
  type InlineImageSource,
} from '../native/inline-image-resources';
import type { ImageDecodePort } from '@docx-editor.dev/core/store';
import type { TrackedPayloadOptions } from './tracked-stage-payloads';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  TreePackageStore,
  headlessViewOfStore,
  type HeadlessDocumentView,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { createEditPlan, EDIT_PLAN_LIMITS } from '../diff/plan';
import { planTargetDocument, type TargetDocumentOptions } from '../diff/plan-target';
import { targetAstKey } from '../diff/target-key';
import { parseTargetMarkdown } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import type { ReviewViewOptions } from './document-view';
import { reviewGroupRevisionReservations } from './group-allocation';
import { planOwnedProposalRemoval, type OwnedProposalGroup } from './owned-proposal-plan';
import { stageOwnedProposalRemoval } from './owned-proposal-stage';
import {
  verifyOwnedProposalContentProofs,
  type OwnedProposalContentProof,
} from './owned-proposal-content';
import { reconcileRecompositionOwnership } from './recompose-ownership';
import {
  replaceOwnedProposals,
  type OwnedProposalReplacementResult,
} from './replace-owned-proposals';
import { readTrackedStructureInput } from './tracked-structure';
import type { TrackTextContext } from './tracked-inline';
import { readWorkingView } from './working-view';
import { normalizeRecompositionTarget } from './recompose-target';

export interface RecomposeRequest {
  /** The revision whose working document the caller used to prepare the complete target. */
  readonly revision: RevisionToken;
  readonly targetMarkdown: string;
  /** Keep the complete returned ledger across decisions and undo; visibility grants no ownership. */
  readonly ownership: readonly OwnedProposalGroup[];
  /** Content evidence retained with the ownership ledger, including resolved entries. */
  readonly contentProofs?: readonly OwnedProposalContentProof[];
}

export interface RecomposeOptions {
  /** Trusted executor services; resources resolve after planning and before the revision check. */
  readonly images?: {
    readonly resolve: InlineImageResolver;
    readonly decodePort: ImageDecodePort;
    readonly signal?: AbortSignal;
  };
  /** Trusted host layout services. These never supply a replacement canonical document. */
  readonly view?: Omit<ReviewViewOptions, 'pendingProperties'>;
  readonly target?: TargetDocumentOptions;
  /** Exact localized Word fallback labels for changed property and atomic payloads. */
  readonly payloads?: TrackedPayloadOptions;
}

const active = new WeakSet<DocxEditorInstance>();
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/**
 * Replan a complete AI target from the current ownership-relative committed document.
 * Only explicitly owned pending proposals are reversed in a detached native read.
 * Accepted/rejected proposals and manual edits remain current document content.
 * A target is revision-bound: any intervening edit refuses, without retrying stale content.
 * Returns the full ownership ledger, including resolved groups that native undo can restore.
 * The replacement stage executes supported text, format, and paragraph changes in plan order.
 */
export async function recomposeOwnedProposals(
  editor: DocxEditorInstance,
  request: RecomposeRequest,
  context: TrackTextContext,
  options: RecomposeOptions = {}
): Promise<OwnedProposalReplacementResult> {
  if (active.has(editor))
    return fail('A review recomposition is already running.', 'unsupported-operation');
  active.add(editor);
  try {
    const copied = copyBoundedJson(
      { request, context, payloads: options.payloads ?? {} },
      EDIT_PLAN_LIMITS
    );
    if (!copied.ok) return fail(copied.reason);
    const input = copied.value as unknown as {
      request: RecomposeRequest;
      context: TrackTextContext;
      payloads: TrackedPayloadOptions;
    };
    if (
      !input.request ||
      typeof input.request !== 'object' ||
      typeof input.request.targetMarkdown !== 'string' ||
      !Array.isArray(input.request.ownership) ||
      Object.keys(input.request).some(
        (key) => !['revision', 'targetMarkdown', 'ownership', 'contentProofs'].includes(key)
      )
    )
      return fail('A recomposition needs a revision, complete target, and ownership ledger.');
    const token = createEditPlan(input.request.revision, []);
    if (!token.success) return token;
    const checked = readTrackedStructureInput(editor, token.plan, input.context, 'recomposition');
    if (!checked.success) return checked;
    const { surface, actual, allowGroupInitialization } = checked.input;
    const pkg = surface.session.currentPackage();
    const fresh = () => {
      if (editor.surface !== surface)
        return fail(
          'The native document changed while recomposition was prepared.',
          'document-not-ready'
        );
      const current = readEditorRevision(editor, input.context.documentId, input.context.epoch);
      if (!current || !sameRevision(current, actual) || surface.session.currentPackage() !== pkg)
        return {
          ...fail(
            'Read the current working document before recomposing proposals.',
            'stale-revision'
          ),
          expectedRevision: actual,
          ...(current ? { actualRevision: current } : {}),
        };
      return null;
    };
    const selected = reconcileRecompositionOwnership(pkg, actual, input.request.ownership);
    if (!selected.success) return selected;
    const proofs = input.request.contentProofs ?? [];
    const proofRefusal = verifyOwnedProposalContentProofs(pkg, actual, selected.selection, proofs);
    if (proofRefusal) return proofRefusal;
    const groupIds = new Set(input.request.ownership.map((group) => group.groupId));
    if (proofs.some((proof) => !groupIds.has(proof.groupId)))
      return fail('Content evidence names a group outside the ownership ledger.');
    const target = parseTargetMarkdown(input.request.targetMarkdown);
    if (!target.success) return target;
    // Geometry preview and final staging share one immutable resolver snapshot per asset.
    const assetSnapshots = new Map<string, Promise<InlineImageSource>>();
    let snapshotBytes = 0;
    const imageServices = options.images;
    const prepareImages = imageServices
      ? (plan: import('../diff/plan').CanonicalEditPlan) =>
          prepareInlineImageResources(
            plan,
            (reference, signal) => {
              let pending = assetSnapshots.get(reference);
              if (!pending) {
                pending = Promise.resolve(imageServices.resolve(reference, signal)).then(
                  (asset) => {
                    if (
                      !(asset.bytes instanceof Uint8Array) ||
                      asset.bytes.byteLength > 64 * 1024 * 1024 - snapshotBytes
                    )
                      throw Error('Image resources exceed the preparation byte budget.');
                    snapshotBytes += asset.bytes.byteLength;
                    return {
                      bytes: new Uint8Array(asset.bytes),
                      mime: asset.mime,
                      widthPoints: asset.widthPoints,
                      heightPoints: asset.heightPoints,
                    };
                  }
                );
                assetSnapshots.set(reference, pending);
              }
              return pending;
            },
            imageServices.decodePort,
            imageServices.signal
          )
      : undefined;
    const viewOptions = {
      ...options.view,
      tableSyntax: options.view?.tableSyntax ?? ('auto' as const),
    };
    const working = await readWorkingView(surface.session, actual, viewOptions);
    if (!working.success) return working;
    const changedWhileReading = fresh();
    if (changedWhileReading) return changedWhileReading;
    const currentTarget = parseTargetMarkdown(working.view.projection.text);
    const sameAst =
      currentTarget.success &&
      targetAstKey(currentTarget.target.normalizedAst) ===
        targetAstKey(target.target.normalizedAst);
    // Projection spelling can differ while every native operation is already satisfied.
    // Check the actual working target before inverses can create a redundant replacement.
    const sameTarget =
      sameAst ||
      isCurrentNativeTarget(
        { revision: actual, pkg, projection: working.view.projection },
        input.request.targetMarkdown,
        options.target?.definitions
      );
    const changedByCurrentPlan = fresh();
    if (changedByCurrentPlan) return changedByCurrentPlan;
    const hasMergedCells = working.view.projection.cells.some(
      (cell) => cell.gridSpan > 1 || cell.verticalContinuation
    );
    const normalized =
      hasMergedCells && !sameTarget
        ? await normalizeRecompositionTarget(
            pkg,
            actual,
            target.target,
            viewOptions,
            options.target,
            fresh,
            prepareImages
          )
        : { success: true as const, markdown: input.request.targetMarkdown, unchanged: false };
    if (!normalized.success) return normalized;
    if (normalized.unchanged || sameTarget)
      return {
        success: true,
        committed: false,
        revision: Object.freeze({ ...actual }),
        ownership: Object.freeze(input.request.ownership.map(freezeOwnership)),
        contentProofs: Object.freeze(proofs.map((proof) => Object.freeze({ ...proof }))),
      };
    // Both planner calls run before observed live capture. Their detached inverse previews
    // must never become primitives in the live collaboration journal.
    const removal = planOwnedProposalRemoval({ pkg, revision: actual }, selected.selection);
    if (!removal.success) return removal;
    const store = new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!, {
      reservedRevisionIds: reviewGroupRevisionReservations,
    });
    const staged = store.transact(
      { kind: 'body' },
      (ctx) =>
        stageOwnedProposalRemoval(ctx, removal.plan, allowGroupInitialization, (run) => {
          if (!run()) throw new Error('A detached owned proposal inverse refused.');
        }),
      { recordsHistory: false, actorId: surface.collaborationSession()?.identity.actorId }
    );
    if (!staged.ok)
      return fail(
        `Unable to prepare the recomposition baseline: ${staged.reason}.`,
        'apply-failed'
      );
    const baseline = store.currentPackage();
    const view: HeadlessDocumentView = {
      ...headlessViewOfStore(store),
      // This is an immutable read of the captured live revision, not its detached edit counter.
      packageRevision: () => actual.revision,
    };
    // Foreign proposals remain proposed. A global committed view would hide content the
    // caller still sees, and would grant ownership over imported or same-author proposals.
    const projected = await readWorkingView(view, actual, viewOptions);
    if (!projected.success) return projected;
    const changed = fresh();
    if (changed) return changed;
    const planned = planTargetDocument(
      { revision: actual, pkg: baseline, projection: projected.view.projection },
      normalized.markdown,
      options.target
    );
    if (!planned.success) return planned;
    const preparedImages = prepareImages ? await prepareImages(planned.plan) : undefined;
    if (preparedImages && !preparedImages.success) return preparedImages;
    const changedByPlanner = fresh();
    if (changedByPlanner) return changedByPlanner;
    const changedContent = verifyOwnedProposalContentProofs(
      pkg,
      actual,
      selected.selection,
      proofs
    );
    if (changedContent) return changedContent;
    const result = replaceOwnedProposals(
      editor,
      selected.selection,
      planned.plan,
      input.context,
      input.payloads,
      preparedImages?.images
    );
    if (!result.success) return result;
    const ledger = new Map(input.request.ownership.map((group) => [group.groupId, group]));
    for (const group of result.ownership) ledger.set(group.groupId, group);
    const contentProofs = new Map(proofs.map((proof) => [proof.groupId, proof]));
    for (const proof of result.contentProofs) contentProofs.set(proof.groupId, proof);
    return {
      ...result,
      ownership: Object.freeze([...ledger.values()].map(freezeOwnership)),
      contentProofs: Object.freeze(
        [...contentProofs.values()].map((proof) => Object.freeze({ ...proof }))
      ),
    };
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'apply-failed');
  } finally {
    active.delete(editor);
  }
}

function freezeOwnership(group: OwnedProposalGroup): OwnedProposalGroup {
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
