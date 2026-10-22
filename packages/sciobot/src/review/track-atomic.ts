import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { RevisionToken } from '../contracts/document';
import type { CanonicalEditPlan } from '../diff/plan';
import { decideApplicationPayloadProposal, type PropertyDecisionResult } from './decide-properties';
import {
  trackApplicationPayloadPlan,
  type PropertyFallbackTarget,
  type TrackPropertiesContext,
  type TrackPropertiesResult,
} from './track-properties';

/**
 * Retain complete old/proposed atomic payloads, including opaque data and asset references.
 * The application supplies the localized Word text fallback. New proposals require one native
 * fallback paragraph per atomic block. The original payload remains canonical until acceptance.
 * Group composition and overlapping later waves belong to R14/R19.
 */
export function trackAtomicPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  fallbackTargets: readonly PropertyFallbackTarget[],
  context: TrackPropertiesContext
): TrackPropertiesResult {
  return trackApplicationPayloadPlan(editor, input, fallbackTargets, context, 'atomic');
}

/** Decide the retained payload and its owned native fallback revisions in one history unit. */
export function decideAtomicProposal(
  editor: DocxEditorInstance,
  proposalId: string,
  decision: 'accept' | 'reject',
  expectedRevision: RevisionToken,
  identity: Pick<TrackPropertiesContext, 'documentId' | 'epoch'>
): PropertyDecisionResult {
  return decideApplicationPayloadProposal(
    editor,
    proposalId,
    decision,
    expectedRevision,
    identity,
    'atomic'
  );
}
