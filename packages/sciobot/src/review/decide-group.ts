import {
  propertyHistoryOf,
  readPropertyGroupReceipt,
  combinePropertyGroupReceipts,
} from './property-history';
import type { PropertyReviewReceipt } from '../collaboration/property-review-prototype';
import { preparePresenceDecision } from './prepare-presence-decision';
import { decodeBlockPresenceRecord } from './block-presence-schema';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  withPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { DEFAULT_PAYLOAD_LIMITS } from '../blocks/payload-schema';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { withReviewDecisionGuard } from './decision-guard';
import { readEditorRevision } from '../projection/revision-cache';
import { readReviewGroups } from './groups';
import { captureReviewGroupHistory, reviewGroupHistoryRefusal } from './group-history';
import { groupRemovalBudget, partRemovalBudget } from './group-removal-budget';
import { groupReviewFieldBudget, partReviewFieldBudget } from './group-field-budget';
import { readPropertyReviewStore } from './property-record-store';
import { ATOMIC_REVIEW_EXECUTOR } from './property-record-schema';
import { preparePayloadProposalDecision } from './prepare-payload-decision';
import {
  withInitializedReviewGroupStore,
  withoutReviewGroupRecord,
  withPayloadGroupDecisionRecord,
  withReviewGroupRecord,
} from './group-record-store';

export type GroupDecisionResult =
  | { readonly success: true; readonly revision: RevisionToken }
  | NativeEditFailure;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/** Resolve one group, or every current group for null, in one native transaction and undo unit. */
export function decideReviewGroup(
  editor: DocxEditorInstance,
  groupId: string | null,
  expectedRevision: RevisionToken,
  identity: Pick<RevisionToken, 'documentId' | 'epoch'>,
  decision: 'accept' | 'reject'
): GroupDecisionResult {
  return withReviewDecisionGuard(editor, () =>
    commitGroupDecision(editor, groupId, expectedRevision, identity, decision)
  );
}

function commitGroupDecision(
  editor: DocxEditorInstance,
  groupId: string | null,
  expectedRevision: RevisionToken,
  identity: Pick<RevisionToken, 'documentId' | 'epoch'>,
  decision: 'accept' | 'reject'
): GroupDecisionResult {
  const copied = copyBoundedJson(
    { groupId, expectedRevision, identity, decision },
    DEFAULT_PAYLOAD_LIMITS
  );
  if (!copied.ok) return fail(copied.reason);
  const input = copied.value as unknown as {
    groupId: string | null;
    expectedRevision: RevisionToken;
    identity: typeof identity;
    decision: 'accept' | 'reject';
  };
  if (
    (input.decision !== 'accept' && input.decision !== 'reject') ||
    (input.groupId !== null && (!input.groupId || typeof input.groupId !== 'string')) ||
    !input.identity ||
    typeof input.identity.documentId !== 'string' ||
    !input.identity.documentId ||
    typeof input.identity.epoch !== 'string' ||
    !input.identity.epoch ||
    !input.expectedRevision ||
    !Number.isSafeInteger(input.expectedRevision.revision) ||
    input.expectedRevision.revision < 0
  )
    return fail('Invalid group decision or document identity.');
  const surface = editor.surface;
  if (!surface) return fail('The native editor is not attached.', 'document-not-ready');
  const historyRefusal = reviewGroupHistoryRefusal(surface.collaborationSession());
  if (historyRefusal) return fail(historyRefusal, 'unsupported-operation');
  const capability = editor.can({ type: 'setReviewDisplayMode', mode: 'proposed' });
  if (!capability.ok) return fail(String(capability.reason), 'unsupported-operation');
  let result: GroupDecisionResult | undefined;
  try {
    const refused = surface.commitReviewOps((historyToken) => {
      const current = readEditorRevision(editor, input.identity.documentId, input.identity.epoch);
      if (!current || editor.surface !== surface) {
        result = fail('The native editor is not attached.', 'document-not-ready');
        return { committed: false, reason: result.error };
      }
      if (!sameRevision(input.expectedRevision, current)) {
        result = {
          success: false,
          code: 'stale-revision',
          error: `Read the current document before ${input.decision === 'accept' ? 'accepting' : 'rejecting'} a review group.`,
          expectedRevision: input.expectedRevision,
          actualRevision: current,
        };
        return { committed: false, reason: result.error };
      }
      const pkg = surface.session.currentPackage(),
        read = readReviewGroups(pkg);
      if (!read.success) {
        result = read;
        return { committed: false, reason: read.error };
      }
      const selected =
        input.groupId === null
          ? read.groups
          : read.groups.filter((group) => group.id === input.groupId);
      if (!selected.length && input.groupId === null) {
        result = { success: true, revision: current };
        return { committed: false };
      }
      if (!selected.length || selected.some((group) => group.readOnly)) {
        result = fail(
          'The current review group is missing or needs reconciliation.',
          'unsupported-operation'
        );
        return { committed: false, reason: result.error };
      }
      const propertyHistory = propertyHistoryOf(surface.collaborationSession());
      const propertyReceipts: PropertyReviewReceipt[] = [];
      const withdrawnGroups = new Set<string>();
      if (input.decision === 'reject' && propertyHistory) {
        for (const group of selected) {
          const receipt = readPropertyGroupReceipt(propertyHistory, group);
          if (receipt) {
            propertyReceipts.push(receipt);
            withdrawnGroups.add(group.id);
          }
        }
      }
      const selectedSites = new Map<string, Set<string>>();
      for (const member of selected
        .filter((group) => !withdrawnGroups.has(group.id))
        .flatMap((group) => group.native)) {
        const ids = selectedSites.get(member.partName) ?? new Set<string>();
        for (const id of member.siteNodeIds) ids.add(id);
        selectedSites.set(member.partName, ids);
      }
      // Structural members can depend on one another. Resolve the exact selected set together.
      const partOps = new Map<string, TreeDocOp[]>(
        [...selectedSites].map(([partName, ids]) => [
          partName,
          [
            {
              op: input.decision === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
              siteNodeIds: [...ids],
            },
          ],
        ])
      );
      const edits: ((pkg: OoxmlPackage) => OoxmlPackage)[] = [];
      for (const groupId of withdrawnGroups) {
        if (!groupId.startsWith('g1:stored:'))
          throw new Error('Property ownership needs a stored native group.');
        const id = groupId.slice('g1:stored:'.length);
        edits.push((working) => {
          const removed = withoutReviewGroupRecord(working, id);
          if (!removed.ok) throw new Error(removed.reason);
          return removed.pkg;
        });
      }
      const proposals = readPropertyReviewStore(pkg);
      for (const group of selected) {
        const expectedPropertySites = new Map<string, Set<string>>();
        for (const member of group.properties) {
          const entry =
            proposals.status === 'ready' ? proposals.records.get(member.proposalId) : undefined;
          if (!entry) {
            result = fail('The current application proposal is missing.', 'invalid-document');
            return { committed: false, reason: result.error };
          }
          if (entry.decoded.status !== 'supported') {
            if (decodeBlockPresenceRecord(entry.raw).status !== 'supported') {
              result = fail('Unsupported application proposal version.', 'unsupported-operation');
              return { committed: false, reason: result.error };
            }
            const presence = preparePresenceDecision(
              pkg,
              member.proposalId,
              member.blockId,
              input.decision
            );
            if (!presence.success) {
              result = presence;
              return { committed: false, reason: result.error };
            }
            if (group.native.length && !presence.nativeSiteNodeIds) {
              result = fail('Presence ownership cannot absorb native revisions.');
              return { committed: false, reason: result.error };
            }
            if (presence.nativeSiteNodeIds) {
              const expected = expectedPropertySites.get(presence.partName) ?? new Set<string>();
              for (const id of presence.nativeSiteNodeIds) expected.add(id);
              expectedPropertySites.set(presence.partName, expected);
            }
            partOps.set(presence.partName, [
              ...(presence.beforeOps ?? []),
              ...(partOps.get(presence.partName) ?? []),
              ...presence.ops,
            ]);
            edits.push(presence.apply);
            continue;
          }
          const prepared = preparePayloadProposalDecision(
            pkg,
            current,
            member.proposalId,
            input.decision,
            entry.decoded.value.extension?.executor === ATOMIC_REVIEW_EXECUTOR
              ? 'atomic'
              : 'properties',
            new Set([member.blockId])
          );
          if (!prepared.success) {
            result = prepared;
            return { committed: false, reason: prepared.error };
          }
          const expectedSites = expectedPropertySites.get(prepared.partName) ?? new Set<string>();
          for (const op of prepared.ops)
            if (op.op === (input.decision === 'accept' ? 'acceptRevision' : 'rejectRevision'))
              for (const id of op.siteNodeIds ?? []) expectedSites.add(id);
          expectedPropertySites.set(prepared.partName, expectedSites);
          edits.push(prepared.apply);
        }
        for (const [partName, expectedSites] of expectedPropertySites) {
          const actualSites = new Set(
            group.native
              .filter((row) => row.partName === partName)
              .flatMap((row) => row.siteNodeIds)
          );
          if (
            expectedSites.size !== actualSites.size ||
            [...expectedSites].some((id) => !actualSites.has(id))
          ) {
            result = fail(
              'The property group has incomplete native membership.',
              'invalid-document'
            );
            return { committed: false, reason: result.error };
          }
        }
        const importedNative = !group.properties.length && !group.id.startsWith('g1:stored:');
        if (group.properties.length || importedNative) {
          const receipts = group.properties.map((member) => ({
            schemaVersion: 2 as const,
            id: `rg_${crypto.randomUUID().replaceAll('-', '')}`,
            ...member,
            decision: input.decision,
          }));
          const nativeRecordId = `rg_${crypto.randomUUID().replaceAll('-', '')}`;
          edits.push((working) => {
            let next = working;
            if (editor.snapshot().collaborationStatus === 'inactive') {
              const initialized = withInitializedReviewGroupStore(next);
              if (!initialized.ok) throw new Error(initialized.reason);
              next = initialized.pkg;
            }
            for (const receipt of receipts) {
              const written = withPayloadGroupDecisionRecord(next, receipt);
              if (!written.ok) throw new Error(written.reason);
              next = written.pkg;
            }
            if (importedNative) {
              const written = withReviewGroupRecord(next, {
                schemaVersion: 1,
                id: nativeRecordId,
                members: group.native.map((member) => ({
                  partName: member.partName,
                  revision: member.address,
                })),
              });
              if (!written.ok) throw new Error(written.reason);
              next = written.pkg;
            }
            return next;
          });
        }
      }
      const collaboration = surface.collaborationSession();
      let previewPackage = pkg,
        flagWrites = 0,
        fieldWrites = 0;
      for (const [partName, ops] of partOps) {
        let preview = pkg.parts.get(partName);
        if (!preview) {
          result = fail('A review part is missing.', 'invalid-document');
          return { committed: false, reason: result.error };
        }
        for (const op of ops) {
          const applied = applyTreeOp(preview, op);
          if (!applied.ok) {
            result = fail(String(applied.reason), 'unsupported-operation');
            return { committed: false, reason: result.error };
          }
          if (collaboration) {
            flagWrites += partRemovalBudget(preview, applied.part);
            fieldWrites += partReviewFieldBudget(preview, applied.part);
          }
          preview = applied.part;
        }
        previewPackage = withPart(previewPackage, preview);
      }
      if (collaboration)
        for (const edit of edits) {
          const next = edit(previewPackage);
          flagWrites += groupRemovalBudget(previewPackage, next);
          fieldWrites += groupReviewFieldBudget(previewPackage, next);
          previewPackage = next;
        }
      const commit = () =>
        captureReviewGroupHistory(
          collaboration,
          () =>
            surface.session.applyTreeOps(
              [],
              undefined,
              undefined,
              { kind: 'body' },
              {
                historyToken,
                partOps: [...partOps].map(([partName, ops]) => ({ partName, ops })),
                packageEditsAfterOps: edits,
              }
            ),
          flagWrites,
          fieldWrites
        );
      const applied =
        propertyReceipts.length && propertyHistory
          ? propertyHistory.withdrawDuringPublication(
              combinePropertyGroupReceipts(propertyReceipts),
              commit,
              flagWrites,
              fieldWrites
            )
          : commit();
      result = applied.committed
        ? { success: true, revision: { ...current, revision: surface.session.packageRevision() } }
        : fail(String(applied.reason ?? 'The group decision did not commit.'), 'apply-failed');
      return applied;
    }, 'revision-resolve');
    return (
      result ??
      fail(String(refused ?? 'The group decision did not commit.'), 'unsupported-operation')
    );
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'apply-failed');
  }
}
