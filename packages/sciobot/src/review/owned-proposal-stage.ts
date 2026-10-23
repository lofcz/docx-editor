import { decodeBlockPresenceRecord } from './block-presence-schema';
import { readPropertyReviewStore } from './property-record-store';
import { preparePresenceDecision } from './prepare-presence-decision';
import type { TreeTransactionContext } from '@docx-editor.dev/core/store';
import type { OwnedProposalRemovalPlan } from './owned-proposal-plan';
import {
  withInitializedReviewGroupStore,
  withPayloadGroupDecisionRecord,
  withReviewGroupRecord,
} from './group-record-store';

/** Apply an already validated inverse through the caller's native transaction and budget gate. */
export function stageOwnedProposalRemoval(
  ctx: TreeTransactionContext,
  removal: OwnedProposalRemovalPlan,
  allowInitialization: boolean,
  step: (run: () => boolean) => void
): void {
  // Capture ownership before native inverses remove the controls. The returned completions
  // retire metadata only after those exact native operations have succeeded.
  const pkg = ctx.currentPackage(),
    records = readPropertyReviewStore(pkg);
  const presence = removal.propertyDecisions.flatMap((member) => {
    const entry = records.status === 'ready' ? records.records.get(member.proposalId) : undefined;
    if (!entry) throw new Error('An owned application inverse record is missing.');
    if (entry.decoded.status === 'supported') return [];
    if (decodeBlockPresenceRecord(entry.raw).status !== 'supported')
      throw new Error('An owned application inverse record is unsupported.');
    const prepared = preparePresenceDecision(pkg, member.proposalId, member.blockId, 'reject');
    if (!prepared.success) throw new Error(prepared.error);
    return [prepared];
  });
  for (const part of removal.partOps)
    for (const op of part.ops) step(() => ctx.applyTo(part.partName, op));
  for (const prepared of presence) step(() => ctx.applyPackage(prepared.apply));
  if (!removal.propertyDecisions.length && !removal.retainNativeGroups.length) return;
  if (allowInitialization)
    step(() =>
      ctx.applyPackage((working) => {
        const initialized = withInitializedReviewGroupStore(working);
        if (!initialized.ok) throw new Error(initialized.reason);
        return initialized.pkg;
      })
    );
  for (const decision of removal.propertyDecisions)
    step(() =>
      ctx.applyPackage((working) => {
        const written = withPayloadGroupDecisionRecord(working, {
          schemaVersion: 2,
          id: `rg_${crypto.randomUUID().replaceAll('-', '')}`,
          ...decision,
        });
        if (!written.ok) throw new Error(written.reason);
        return written.pkg;
      })
    );
  for (const group of removal.retainNativeGroups)
    step(() =>
      ctx.applyPackage((working) => {
        const written = withReviewGroupRecord(working, {
          schemaVersion: 1,
          id: `rg_${crypto.randomUUID().replaceAll('-', '')}`,
          members: group.members,
        });
        if (!written.ok) throw new Error(written.reason);
        return written.pkg;
      })
    );
}
