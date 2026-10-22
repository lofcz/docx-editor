import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { DEFAULT_PAYLOAD_LIMITS, type BlockPayload } from '../blocks/payload-schema';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { readEditorRevision } from '../projection/revision-cache';
import { withReviewDecisionGuard } from './decision-guard';
import {
  prepareExternalReview,
  type ExternalReviewReport,
  type ExternalReviewTarget,
} from './external-review-state';
import { captureReviewGroupHistory, reviewGroupHistoryRefusal } from './group-history';
import { groupRemovalBudget } from './group-removal-budget';
import { groupReviewFieldBudget } from './group-field-budget';
import {
  withInitializedReviewGroupStore,
  withPayloadGroupDecisionRecord,
} from './group-record-store';

export { inspectExternalReview } from './external-review-state';
export type {
  ExternalReviewEntry,
  ExternalReviewStatus,
  ExternalReviewConflict,
  ExternalReviewReport,
  ExternalReviewRead,
} from './external-review-state';
export type ReconcileExternalReviewResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly report: ExternalReviewReport;
      readonly reconciled: readonly {
        readonly proposalId: string;
        readonly blockId: string;
        readonly decision: 'accept' | 'reject';
      }[];
    }
  | NativeEditFailure;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/** Metadata only. The already-reviewed Word content is never rewritten. */
function acceptPayload(pkg: OoxmlPackage, target: ExternalReviewTarget): OoxmlPackage {
  const owner = buildBlockIndex(pkg).byId.get(target.block.blockId);
  if (!owner) throw new Error('The external review owner disappeared.');
  const changed = new Set(target.block.changes.map((change) => change.name));
  const payload = target.atomic
    ? target.block.changes[0]!.after.present
      ? (target.block.changes[0]!.after.value as BlockPayload)
      : null
    : {
        ...owner.payload,
        attributes: Object.fromEntries([
          ...Object.entries(owner.payload.attributes).filter(([key]) => !changed.has(key)),
          ...target.block.changes.flatMap((change) =>
            change.after.present ? [[change.name, change.after.value]] : []
          ),
        ]),
      };
  if (!payload) throw new Error('An atomic reconciliation needs its complete proposed payload.');
  const written = withPayloadRecord(pkg, payload, target.label.after);
  if (!written.ok) throw new Error(written.reason);
  return written.pkg;
}

/** Reconcile clear external decisions in one native history frame; keep every ambiguous block. */
export function reconcileExternalReview(
  editor: DocxEditorInstance,
  expectedRevision: RevisionToken,
  identity: Pick<RevisionToken, 'documentId' | 'epoch'>
): ReconcileExternalReviewResult {
  return withReviewDecisionGuard(editor, () => reconcile(editor, expectedRevision, identity));
}

function reconcile(
  editor: DocxEditorInstance,
  expectedRevision: RevisionToken,
  identity: Pick<RevisionToken, 'documentId' | 'epoch'>
): ReconcileExternalReviewResult {
  const copied = copyBoundedJson({ expectedRevision, identity }, DEFAULT_PAYLOAD_LIMITS);
  if (!copied.ok) return fail(copied.reason);
  const input = copied.value as unknown as {
    expectedRevision: RevisionToken;
    identity: typeof identity;
  };
  if (
    !input.identity ||
    typeof input.identity.documentId !== 'string' ||
    !input.identity.documentId.trim() ||
    typeof input.identity.epoch !== 'string' ||
    !input.identity.epoch.trim() ||
    !input.expectedRevision ||
    !Number.isSafeInteger(input.expectedRevision.revision) ||
    input.expectedRevision.revision < 0
  )
    return fail('Invalid external review identity or revision.');
  const surface = editor.surface;
  if (!surface) return fail('The native editor is not attached.', 'document-not-ready');
  const history = reviewGroupHistoryRefusal(surface.collaborationSession());
  if (history) return fail(history, 'unsupported-operation');
  const capability = editor.can({ type: 'setReviewDisplayMode', mode: 'proposed' });
  if (!capability.ok) return fail(String(capability.reason), 'unsupported-operation');
  let result: ReconcileExternalReviewResult | undefined;
  try {
    const refused = surface.commitReviewOps((historyToken) => {
      const current = readEditorRevision(editor, input.identity.documentId, input.identity.epoch);
      if (!current || editor.surface !== surface) {
        result = fail('The native editor is not attached.', 'document-not-ready');
        return { committed: false, reason: result.error };
      }
      if (!sameRevision(current, input.expectedRevision)) {
        result = {
          success: false,
          code: 'stale-revision',
          error: 'Read the current document before reconciling external review.',
          expectedRevision: input.expectedRevision,
          actualRevision: current,
        };
        return { committed: false, reason: result.error };
      }
      const pkg = surface.session.currentPackage(),
        prepared = prepareExternalReview(pkg);
      if (!prepared.success) {
        result = prepared;
        return { committed: false, reason: prepared.error };
      }
      const reconciled = Object.freeze(
        prepared.targets.map((target) =>
          Object.freeze({
            proposalId: target.record.id,
            blockId: target.block.blockId,
            decision: target.decision,
          })
        )
      );
      if (!reconciled.length) {
        result = Object.freeze({
          success: true,
          revision: Object.freeze(current),
          report: prepared.report,
          reconciled,
        });
        return { committed: false };
      }
      const edits: ((working: OoxmlPackage) => OoxmlPackage)[] = [];
      if (editor.snapshot().collaborationStatus === 'inactive')
        edits.push((working) => {
          const initialized = withInitializedReviewGroupStore(working);
          if (!initialized.ok) throw new Error(initialized.reason);
          return initialized.pkg;
        });
      for (const target of prepared.targets) {
        if (target.decision === 'accept') edits.push((working) => acceptPayload(working, target));
        const receipt = Object.freeze({
          schemaVersion: 2 as const,
          id: `rg_${crypto.randomUUID().replaceAll('-', '')}`,
          proposalId: target.record.id,
          blockId: target.block.blockId,
          decision: target.decision,
        });
        edits.push((working) => {
          const written = withPayloadGroupDecisionRecord(working, receipt);
          if (!written.ok) throw new Error(written.reason);
          return written.pkg;
        });
      }
      const collaboration = surface.collaborationSession();
      let preview = pkg,
        flagWrites = 0,
        fieldWrites = 0;
      for (const edit of edits) {
        const next = edit(preview);
        if (collaboration) {
          flagWrites += groupRemovalBudget(preview, next);
          fieldWrites += groupReviewFieldBudget(preview, next);
        }
        preview = next;
      }
      const applied = captureReviewGroupHistory(
        collaboration,
        () =>
          surface.session.applyTreeOps(
            [],
            undefined,
            undefined,
            { kind: 'body' },
            { historyToken, packageEditsAfterOps: edits }
          ),
        flagWrites,
        fieldWrites
      );
      result = applied.committed
        ? Object.freeze({
            success: true,
            revision: Object.freeze({ ...current, revision: surface.session.packageRevision() }),
            report: prepared.report,
            reconciled,
          })
        : fail(
            String(applied.reason ?? 'External review reconciliation did not commit.'),
            'apply-failed'
          );
      return applied;
    }, 'revision-resolve');
    return (
      result ??
      fail(
        String(refused ?? 'External review reconciliation did not commit.'),
        'unsupported-operation'
      )
    );
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'apply-failed');
  }
}
