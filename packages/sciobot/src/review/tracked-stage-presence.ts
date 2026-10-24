import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { CanonicalEditPlan } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
import { createBlockPresenceRecord } from './block-presence-plan';
import { withBlockPresenceRecord } from './block-presence-store';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

/** Capture complete, fully initialized application subtrees after all ordered native stages. */
export function stageTrackedPresence(
  writer: TrackedStageWriter,
  before: OoxmlPackage,
  plan: CanonicalEditPlan,
  attribution: TrackedStageAttribution,
  supplied?: { readonly proposalId: string }
): void {
  const births = plan.operations.flatMap((edit) =>
    edit.kind === 'inline-atoms'
      ? edit.atoms.flatMap((atom) => (atom.kind === 'inline-control' ? [atom.payload.id] : []))
      : []
  );
  const removals = plan.operations.filter((edit) => edit.kind === 'inline-remove');
  const roots = plan.operations.flatMap((edit) =>
    edit.kind === 'siblings'
      ? (edit.blocks ?? [])
          .filter((binding) => binding.path.length === 1)
          .map((binding) => binding.payload.id)
      : []
  );
  if (!roots.length && !removals.length && !births.length) {
    if (supplied)
      refuseWrite('Presence ownership has no inserted application roots.', 'invalid-input');
    return;
  }
  if (
    !supplied ||
    typeof supplied !== 'object' ||
    Array.isArray(supplied) ||
    Object.keys(supplied).some((key) => key !== 'proposalId')
  )
    refuseWrite(
      'Inserted application blocks need an explicit presence proposal identity.',
      'unsupported-operation'
    );
  const prepared = createBlockPresenceRecord(
    before,
    writer.currentPackage(),
    [
      ...[...roots, ...births].map((blockId) => ({
        blockId,
        before: { present: false },
        after: { present: true },
      })),
      ...removals.map((edit) => ({
        blockId: edit.blockId,
        before: { present: true },
        after: { present: false },
      })),
    ],
    { id: supplied.proposalId, ...attribution },
    removals.length || births.length ? { nativeInline: true } : {}
  );
  if (!prepared.success) refuseWrite(prepared.error, prepared.code);
  writer.edit((pkg) => {
    const stored = withBlockPresenceRecord(pkg, prepared.record);
    if (!stored.success) refuseWrite(stored.error, stored.code);
    return stored.pkg;
  });
}
