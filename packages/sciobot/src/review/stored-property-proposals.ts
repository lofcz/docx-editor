import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { createEditPlan, type CanonicalEdit } from '../diff/plan';
import { createPendingPropertyProposal, type PendingPropertyProposal } from './property-projection';
import { type StoredPropertyRecord, readPropertyReviewStore } from './property-record-store';
import { ATOMIC_PAYLOAD_CHANGE, ATOMIC_REVIEW_EXECUTOR } from './property-record-schema';
import { readPayloadGroupDecisions, payloadGroupDecisionKey } from './payload-group-decisions';

/** Bind persisted stable block IDs to this load's actual canonical owners. Never reuse saved node IDs. */
export function readStoredPropertyProposals(
  pkg: OoxmlPackage,
  revision: RevisionToken
):
  | { readonly success: true; readonly pending: readonly PendingPropertyProposal[] }
  | NativeEditFailure {
  const store = readPropertyReviewStore(pkg);
  if (store.status === 'absent') return { success: true, pending: Object.freeze([]) };
  if (store.status === 'invalid')
    return { success: false, code: 'invalid-document', error: store.reason };
  return bindStoredPropertyProposals(pkg, revision, store.records);
}
/** Bind an admitted version-one subset. Combined readers retain the complete causal ancestry. */
export function bindStoredPropertyProposals(
  pkg: OoxmlPackage,
  revision: RevisionToken,
  records: ReadonlyMap<string, StoredPropertyRecord>
): ReturnType<typeof readStoredPropertyProposals> {
  const decisions = readPayloadGroupDecisions(pkg);
  if (decisions.status === 'invalid')
    return { success: false, code: 'invalid-document', error: decisions.reason };
  const index = buildBlockIndex(pkg),
    pending: PendingPropertyProposal[] = [],
    remaining = new Map(records),
    ancestors = new Map<string, Set<string>>(),
    lastProperty = new Map<string, string>();
  // Node order and wall clocks cannot establish proposal causality after concurrent materialization.
  while (remaining.size) {
    for (const entry of remaining.values())
      if (entry.decoded.status === 'unsupported-version')
        return {
          success: false,
          code: 'unsupported-operation',
          error: `Unsupported property proposal version ${entry.decoded.version}.`,
        };
    const ready = [...remaining.values()]
      .filter(
        (entry) =>
          entry.decoded.status === 'supported' &&
          entry.decoded.value.predecessors.every((id) => !remaining.has(id))
      )
      .sort((a, b) => a.id.localeCompare(b.id));
    if (!ready.length)
      return {
        success: false,
        code: 'invalid-document',
        error: 'Property proposal predecessors contain a cycle.',
      };
    for (const entry of ready) {
      if (entry.decoded.status === 'unsupported-version')
        return {
          success: false,
          code: 'unsupported-operation',
          error: `Unsupported property proposal version ${entry.decoded.version}.`,
        };
      const operations: CanonicalEdit[] = [];
      const prior = new Set(entry.decoded.value.predecessors);
      for (const id of entry.decoded.value.predecessors)
        for (const ancestor of ancestors.get(id) ?? []) prior.add(ancestor);
      for (const saved of entry.decoded.value.blocks) {
        const disposition = decisions.decisions.get(
          payloadGroupDecisionKey(entry.id, saved.blockId)
        );
        if (disposition?.status === 'conflict')
          return {
            success: false,
            code: 'ambiguous-target',
            error: 'Conflicting application group decisions require reconciliation.',
          };
        if (disposition) continue;
        const atomic = entry.decoded.value.extension?.executor === ATOMIC_REVIEW_EXECUTOR;
        if (
          atomic &&
          (saved.changes.length !== 1 || saved.changes[0]!.name !== ATOMIC_PAYLOAD_CHANGE)
        )
          return {
            success: false,
            code: 'invalid-document',
            error: 'Atomic proposals require one complete payload pair per block.',
          };
        const block = index.byId.get(saved.blockId);
        if (!block || block.payload.type !== saved.type)
          return {
            success: false,
            code: 'invalid-document',
            error: `Property proposal ${entry.id} has no unique native ${saved.type} owner for ${saved.blockId}.`,
          };
        for (const change of saved.changes) {
          const key = JSON.stringify([saved.blockId, change.name]),
            previous = lastProperty.get(key);
          if (previous && !prior.has(previous))
            return {
              success: false,
              code: 'ambiguous-target',
              error: 'Concurrent property proposals overlap without causal ownership.',
            };
          lastProperty.set(key, entry.id);
          operations.push({
            kind: 'payload',
            blockId: block.id,
            partName: block.partName,
            controlId: block.controlId,
            path: atomic ? [] : ['attributes', change.name],
            before: change.before,
            after: change.after,
          });
        }
      }
      if (operations.length) {
        const plan = createEditPlan(revision, operations);
        if (!plan.success) return plan;
        const proposal = createPendingPropertyProposal(entry.id, plan.plan);
        if (!proposal.success) return proposal;
        pending.push(proposal.proposal);
      }
      ancestors.set(entry.id, prior);
      remaining.delete(entry.id);
    }
  }
  return { success: true, pending: Object.freeze(pending) };
}
