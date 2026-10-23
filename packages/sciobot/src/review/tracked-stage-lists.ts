import { collectRevisionSites, type RevisionAddress } from '@docx-editor.dev/core/store';
import { applyDirectListAssignments, type DirectListContext } from '../ai/direct-lists';
import { planListMembership } from '../diff/lists';
import type { CanonicalEdit, CanonicalEditPlan } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
import type { ReviewGroupReceipt } from './groups';
import {
  nativePropertyAddress,
  propertyControlRefusal,
  propertyAddressKey,
} from './property-fallback';
import { stageTrackedChecklist } from './tracked-stage-checklists';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

type Edit = Extract<CanonicalEdit, { kind: 'list' }>;
/** Preserve original paragraph numbering inside native property revisions. */
export function stageTrackedList(
  writer: TrackedStageWriter,
  plan: CanonicalEditPlan,
  edit: Edit,
  revision: TrackedStageAttribution,
  context: DirectListContext
): readonly ReviewGroupReceipt[] {
  if (edit.after.kind === 'checklist' || edit.before.checkbox)
    return stageTrackedChecklist(writer, plan, edit, revision, context);
  const pkg = writer.currentPackage(),
    part = pkg.parts.get(edit.partName)!;
  const protection = propertyControlRefusal(part, edit.paragraphId);
  if (protection) refuseWrite(protection, 'unsupported-operation');
  const checked = planListMembership({ pkg, partName: edit.partName, revision: plan.revision }, [
    { paragraphId: edit.paragraphId, target: edit.after },
  ]);
  if (!checked.success) refuseWrite(checked.error, checked.code);
  const current = checked.plan.operations[0];
  if (
    !current ||
    current.kind !== 'list' ||
    JSON.stringify(current.before) !== JSON.stringify(edit.before)
  )
    refuseWrite('The list source changed after planning.', 'stale-revision');
  const originalSites = new Set(collectRevisionSites(part).map((site) => site.node.id));
  writer.edit((working) =>
    applyDirectListAssignments(
      working,
      edit.partName,
      [{ paragraphId: edit.paragraphId, before: edit.before, target: edit.after }],
      context,
      revision
    )
  );
  const fresh = collectRevisionSites(writer.currentPackage().parts.get(edit.partName)!).filter(
    (site) => !originalSites.has(site.node.id)
  );
  if (!fresh.length) return [];
  writer.apply(edit.partName, {
    op: 'setRevisionAttribution',
    revision,
    sites: fresh.map((site) => ({ nodeId: site.node.id, renumber: `list:${site.node.localName}` })),
  });
  const addresses = new Map<string, RevisionAddress>();
  for (const site of collectRevisionSites(writer.currentPackage().parts.get(edit.partName)!)) {
    if (originalSites.has(site.node.id)) continue;
    const address = nativePropertyAddress(site.node);
    if (!address || site.refused)
      refuseWrite('The list revision cannot be reviewed.', 'unsupported-operation');
    addresses.set(propertyAddressKey(edit.partName, address), address);
  }
  return [{ partName: edit.partName, addresses: [...addresses.values()] }];
}
