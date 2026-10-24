import {
  applyTreeOp,
  TreePackageStore,
  type TreeDocOp,
  collectRevisionSites,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { planListMembership } from '../diff/lists';
import type { RevisionToken } from '../contracts/document';
import { stageTrackedChecklist } from '../review/tracked-stage-checklists';
import { nativePropertyAddress, propertyAddressKey } from '../review/property-fallback';
import { reviewGroupRevisionReservations } from '../review/group-allocation';
import { groupRemovalBudget } from '../review/group-removal-budget';
import { groupReviewFieldBudget } from '../review/group-field-budget';
import { persistReviewGroups, readReviewGroups, type ReviewGroupReceipt } from '../review/groups';
import type { TrackedStageWriter, TrackedStageAttribution } from '../review/tracked-stage-writer';

/** Keep the proposed break, removed selection, and new checkbox in one review decision. */
export function completeTrackedChecklistSplit(
  before: OoxmlPackage,
  split: OoxmlPackage,
  partName: string,
  paragraphId: string,
  checked: boolean,
  level: number,
  revision: RevisionToken,
  attribution: TrackedStageAttribution,
  allowInitialization: boolean,
  measure?: (before: OoxmlPackage, after: OoxmlPackage) => void
): OoxmlPackage {
  let working = split;
  const writer: TrackedStageWriter = {
    currentPackage: () => working,
    apply(name, op) {
      const part = working.parts.get(name);
      if (!part) throw new Error('The checklist story disappeared.');
      const result = applyTreeOp(part, op);
      if (!result.ok) throw new Error(`The tracked checklist refused: ${result.reason}.`);
      const prior = working;
      working = withPart(working, result.part);
      measure?.(prior, working);
    },
    edit(change) {
      const prior = working;
      working = change(working);
      measure?.(prior, working);
    },
  };
  const planned = planListMembership({ pkg: working, partName, revision }, [
    { paragraphId, target: { kind: 'checklist', checked, level } },
  ]);
  if (!planned.success) throw new Error(planned.error);
  const edit = planned.plan.operations[0];
  if (!edit || edit.kind !== 'list')
    throw new Error('The new checklist item has no membership change.');
  stageTrackedChecklist(writer, planned.plan, edit, attribution);
  const prior = new Set(
    collectRevisionSites(before.parts.get(partName)!).flatMap((site) => {
      const address = nativePropertyAddress(site.node);
      return address ? [propertyAddressKey(partName, address)] : [];
    })
  );
  const priorNodes = new Set(
    collectRevisionSites(before.parts.get(partName)!).map((site) => site.node.id)
  );
  const addresses = collectRevisionSites(working.parts.get(partName)!).flatMap((site) => {
    if (priorNodes.has(site.node.id)) return [];
    const address = nativePropertyAddress(site.node);
    if (address && prior.has(propertyAddressKey(partName, address))) return [];
    if (site.refused || !address) throw new Error('The checklist continuation cannot be reviewed.');
    return [address];
  });
  const groups = readReviewGroups(before);
  if (!groups.success) throw new Error(groups.error);
  const inherited: ReviewGroupReceipt[] = groups.groups
    .filter((group) => !group.readOnly)
    .flatMap((group) => {
      const addresses = group.native
        .filter((member) => member.partName === partName)
        .map((member) => member.address);
      return addresses.length ? [{ partName, addresses }] : [];
    });
  const complete = persistReviewGroups(
    before,
    working,
    [...inherited, { partName, addresses }],
    allowInitialization
  );
  measure?.(working, complete);
  return complete;
}

/** A detached native preflight measures every dependent write before shared history admission. */
export function checklistSplitBudget(
  before: OoxmlPackage,
  partName: string,
  ops: readonly TreeDocOp[],
  complete: (
    pkg: OoxmlPackage,
    measure: (before: OoxmlPackage, after: OoxmlPackage) => void
  ) => OoxmlPackage
) {
  let flags = 0,
    fields = 0;
  const measure = (prior: OoxmlPackage, after: OoxmlPackage) => {
    flags += groupRemovalBudget(prior, after);
    fields += groupReviewFieldBudget(prior, after);
  };
  const preview = new TreePackageStore(before, before.parts.get(before.mainDocumentPart)!, {
    reservedRevisionIds: reviewGroupRevisionReservations,
  });
  const result = preview.transact(
    { kind: 'body' },
    (ctx) => {
      for (const op of ops) {
        const prior = ctx.currentPackage();
        if (!ctx.applyTo(partName, op)) throw new Error('The checklist split preflight refused.');
        measure(prior, ctx.currentPackage());
      }
      if (!ctx.applyPackage((pkg) => complete(pkg, measure)))
        throw new Error('The checklist membership preflight refused.');
    },
    { recordsHistory: false }
  );
  if (!result.ok) throw new Error(`The checklist split preflight refused: ${result.reason}.`);
  return { flags, fields };
}
