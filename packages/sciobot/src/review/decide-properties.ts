import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { DEFAULT_PAYLOAD_LIMITS } from '../blocks/payload-schema';
import { copyBoundedJson } from '../data/json';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import type { TrackTextContext } from './tracked-inline';
import { readEditorRevision } from '../projection/revision-cache';
import { withoutPropertyReviewRecord } from './property-record-store';
import { preparePayloadProposalDecision } from './prepare-payload-decision';
import { withReviewDecisionGuard } from './decision-guard';

export type PropertyDecisionResult =
  | { readonly success: true; readonly revision: RevisionToken }
  | NativeEditFailure;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/** Resolve one property's native fallback and metadata in one guarded, separate history transaction. */
export function decidePropertyProposal(
  editor: DocxEditorInstance,
  proposalId: string,
  decision: 'accept' | 'reject',
  expectedRevision: RevisionToken,
  identity: Pick<TrackTextContext, 'documentId' | 'epoch'>
): PropertyDecisionResult {
  return decideApplicationPayloadProposal(
    editor,
    proposalId,
    decision,
    expectedRevision,
    identity,
    'properties'
  );
}

/** Internal shared decision. Complete atomic pairs never become attribute writes. */
export function decideApplicationPayloadProposal(
  editor: DocxEditorInstance,
  proposalId: string,
  decision: 'accept' | 'reject',
  expectedRevision: RevisionToken,
  identity: Pick<TrackTextContext, 'documentId' | 'epoch'>,
  executor: 'properties' | 'atomic'
): PropertyDecisionResult {
  return withReviewDecisionGuard(editor, () =>
    commitApplicationPayloadDecision(
      editor,
      proposalId,
      decision,
      expectedRevision,
      identity,
      executor
    )
  );
}

function commitApplicationPayloadDecision(
  editor: DocxEditorInstance,
  proposalId: string,
  decision: 'accept' | 'reject',
  expectedRevision: RevisionToken,
  identity: Pick<TrackTextContext, 'documentId' | 'epoch'>,
  executor: 'properties' | 'atomic'
): PropertyDecisionResult {
  const copy = copyBoundedJson(
    { proposalId, decision, expectedRevision, identity },
    DEFAULT_PAYLOAD_LIMITS
  );
  if (!copy.ok) return fail(copy.reason);
  const data = copy.value as unknown as {
      proposalId: string;
      decision: 'accept' | 'reject';
      expectedRevision: RevisionToken;
      identity: Pick<TrackTextContext, 'documentId' | 'epoch'>;
    },
    expected = data.expectedRevision;
  if (
    !expected ||
    !data.identity ||
    typeof data.identity.documentId !== 'string' ||
    !data.identity.documentId ||
    typeof data.identity.epoch !== 'string' ||
    !data.identity.epoch ||
    typeof expected.documentId !== 'string' ||
    typeof expected.epoch !== 'string' ||
    !['accept', 'reject'].includes(data.decision)
  )
    return fail('Invalid property decision or document identity.');
  const surface = editor.surface,
    current = readEditorRevision(editor, data.identity.documentId, data.identity.epoch);
  if (!surface || !current) return fail('The native editor is not attached.', 'document-not-ready');
  if (!sameRevision(expected, current))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read the current document before resolving a property proposal.',
      expectedRevision: expected,
      actualRevision: current,
    };
  if (surface.editingMode() === 'view')
    return fail('The document is open for viewing.', 'unsupported-operation');
  const prepared = preparePayloadProposalDecision(
    surface.session.currentPackage(),
    current,
    data.proposalId,
    data.decision,
    executor
  );
  if (!prepared.success) return prepared;
  try {
    const result = surface.applyAutomationOps(
      () => prepared.ops,
      { kind: 'body' },
      [
        (working) => {
          const applied = prepared.apply(working);
          const removed = withoutPropertyReviewRecord(applied, data.proposalId);
          if (!removed.ok) throw new Error(removed.reason);
          return removed.pkg;
        },
      ],
      true,
      true,
      'after'
    );
    if (result.rejected || !result.committed)
      return fail(String(result.reason ?? 'The property decision did not commit.'), 'apply-failed');
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'apply-failed');
  }
  return { success: true, revision: { ...current, revision: surface.session.packageRevision() } };
}
