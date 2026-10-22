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
  for (const part of removal.partOps)
    for (const op of part.ops) step(() => ctx.applyTo(part.partName, op));
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
