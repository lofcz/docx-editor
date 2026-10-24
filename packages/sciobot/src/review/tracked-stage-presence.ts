import { buildBlockIndex } from '../blocks/block-index';
import { nativePropertyAddress, propertyAddressKey } from './property-fallback';
import type { InlinePresenceNative } from './inline-presence-native';
import type { ReviewGroupReceipt } from './groups';
import { parentNodeOf, type OoxmlPackage } from '@docx-editor.dev/core/store';
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
  supplied?: { readonly proposalId: string },
  receipts: readonly ReviewGroupReceipt[] = []
): void {
  const births = plan.operations.flatMap((edit) =>
    edit.kind === 'inline-atoms' || edit.kind === 'siblings'
      ? (edit.atoms ?? []).flatMap((atom) =>
          atom.kind === 'inline-control' ? [atom.payload.id] : []
        )
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
  const bornInline: InlinePresenceNative[] = [];
  const index = buildBlockIndex(writer.currentPackage());
  for (const edit of plan.operations) {
    if (edit.kind !== 'siblings') continue;
    for (const atom of edit.atoms ?? []) {
      if (atom.kind !== 'inline-control') continue;
      const control = index.byId.get(atom.payload.id);
      const part = control && writer.currentPackage().parts.get(control.partName);
      let parent = part && control ? parentNodeOf(part, control.controlId) : null;
      while (parent && parent.kind !== 'revisionInsert' && parent.kind !== 'paragraph')
        parent = parentNodeOf(part!, parent.id);
      const address = parent?.kind === 'revisionInsert' ? nativePropertyAddress(parent) : null;
      if (!control || !address)
        refuseWrite('The born atom has no paragraph insertion.', 'apply-failed');
      const matching = receipts.filter(
        (receipt) =>
          receipt.partName === control.partName &&
          receipt.addresses.some(
            (candidate) =>
              propertyAddressKey(control.partName, candidate) ===
              propertyAddressKey(control.partName, address)
          )
      );
      if (matching.length !== 1)
        refuseWrite('The born atom has no unique structural receipt.', 'ambiguous-target');
      bornInline.push({
        blockId: atom.payload.id,
        partName: control.partName,
        kind: 'ins',
        revision: address,
        gap: matching[0]!.addresses,
      });
    }
  }
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
    removals.length || births.length
      ? { nativeInline: true, ...(bornInline.length ? { bornInline } : {}) }
      : {}
  );
  if (!prepared.success) refuseWrite(prepared.error, prepared.code);
  writer.edit((pkg) => {
    const stored = withBlockPresenceRecord(pkg, prepared.record);
    if (!stored.success) refuseWrite(stored.error, stored.code);
    return stored.pkg;
  });
}
