import { withReviewGroupRevisionReservations } from './group-allocation';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { applyTreeOp } from '@docx-editor.dev/core/store';
import { DEFAULT_PAYLOAD_LIMITS, isBlockId } from '../blocks/payload-schema';
import { copyBoundedJson } from '../data/json';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { serializeEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import type { PropertyReviewRecord } from './property-record-schema';
import { canonicalTrackedDate, type TrackTextContext } from './tracked-inline';
import { persistReviewGroups } from './groups';
import { prepareApplicationPayloadStage } from './prepare-payload-stage';

/** The application supplies the exact localized label; the adapter never invents content. */
export interface PropertyFallbackTarget {
  readonly blockId: string;
  readonly before: string;
  readonly after: string;
}
export interface TrackPropertiesContext extends TrackTextContext {
  readonly proposalId: string;
}
export type TrackPropertiesResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposal: PropertyReviewRecord | null;
    }
  | NativeEditFailure;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/** Propose attribute deltas and generated Word label edits in one gated native transaction. */
export function trackPropertyPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  fallbackTargets: readonly PropertyFallbackTarget[],
  context: TrackPropertiesContext
): TrackPropertiesResult {
  return trackApplicationPayloadPlan(editor, input, fallbackTargets, context, 'properties');
}

/** Internal shared transaction. The executor selects the stored change semantics. */
export function trackApplicationPayloadPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  fallbackTargets: readonly PropertyFallbackTarget[],
  context: TrackPropertiesContext,
  executor: 'properties' | 'atomic'
): TrackPropertiesResult {
  return withReviewGroupRevisionReservations(editor.surface?.session.currentPackage() ?? null, () =>
    trackApplicationPayloadPlanReserved(editor, input, fallbackTargets, context, executor)
  );
}

function trackApplicationPayloadPlanReserved(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  fallbackTargets: readonly PropertyFallbackTarget[],
  context: TrackPropertiesContext,
  executor: 'properties' | 'atomic'
): TrackPropertiesResult {
  const copied = copyBoundedJson({ context, fallbackTargets }, DEFAULT_PAYLOAD_LIMITS);
  if (!copied.ok) return fail(copied.reason);
  const data = copied.value as unknown as {
      context: TrackPropertiesContext;
      fallbackTargets: PropertyFallbackTarget[];
    },
    attribution = data.context;
  if (
    !attribution ||
    !isBlockId(attribution.proposalId) ||
    typeof attribution.documentId !== 'string' ||
    !attribution.documentId ||
    typeof attribution.epoch !== 'string' ||
    !attribution.epoch ||
    typeof attribution.author !== 'string' ||
    !attribution.author.trim() ||
    !Array.isArray(data.fallbackTargets)
  )
    return fail('Invalid tracked property context or fallback targets.');
  if (typeof attribution.date !== 'string') return fail('Property proposals need a UTC date.');
  const date = canonicalTrackedDate(attribution.date);
  if (!date) return fail('Property proposals need a valid UTC date.');
  let plan: CanonicalEditPlan;
  try {
    plan = JSON.parse(serializeEditPlan(input)) as CanonicalEditPlan;
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause));
  }
  const current = readEditorRevision(editor, attribution.documentId, attribution.epoch),
    surface = editor.surface;
  if (!surface || !current) return fail('The native editor is not attached.', 'document-not-ready');
  if (!sameRevision(current, plan.revision))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read the current document before proposing properties.',
      expectedRevision: plan.revision,
      actualRevision: current,
    };
  if (surface.editingMode() === 'view')
    return fail('The document is open for viewing.', 'unsupported-operation');
  if (!plan.operations.length && !data.fallbackTargets.length)
    return { success: true, revision: current, proposal: null };
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const prepared = prepareApplicationPayloadStage(
    pkg,
    current,
    plan,
    data.fallbackTargets,
    attribution,
    executor
  );
  if (!prepared.success) return prepared;
  const ops = prepared.ops;
  let preview = part;
  for (const op of ops) {
    const result = applyTreeOp(preview, op);
    if (!result.ok)
      return fail(`Native property fallback refused: ${result.reason}.`, 'unsupported-operation');
    preview = result.part;
  }
  let receipt: PropertyReviewRecord | null = null;
  try {
    const result = surface.applyAutomationOps(
      () => ops,
      { kind: 'body' },
      [
        (working) => {
          const complete = prepared.complete(working);
          receipt = complete.receipt;
          return persistReviewGroups(
            pkg,
            complete.pkg,
            [],
            editor.snapshot().collaborationStatus === 'inactive'
          );
        },
      ],
      true,
      true,
      'after'
    );
    if (result.rejected || !result.committed)
      return fail(String(result.reason ?? 'The property proposal did not commit.'), 'apply-failed');
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'apply-failed');
  }
  return {
    success: true,
    revision: { ...current, revision: surface.session.packageRevision() },
    proposal: receipt,
  };
}
