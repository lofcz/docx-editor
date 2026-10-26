import { insertedAtomsOfEdit } from '../diff/inserted-cell-atoms';
import type { InlineImageRuntime } from '../native/inline-image-runtime';
import { buildBlockIndex } from '../blocks/block-index';
import { nativePropertyAddress, propertyAddressKey } from './property-fallback';
import type { InlinePresenceNative } from './inline-presence-native';
import type { ReviewGroupReceipt } from './groups';
import { collectRevisionSites, parentNodeOf, type OoxmlPackage } from '@docx-editor.dev/core/store';
import type { CanonicalEditPlan } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
import { createBlockPresenceRecord } from './block-presence-plan';
import { withBlockPresenceRecord } from './block-presence-store';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

/** Read-only admission for a fresh host presence identity, after target planning. */
export function needsTrackedPresence(plan: CanonicalEditPlan, hasImages = false): boolean {
  return plan.operations.some(
    (edit) =>
      edit.kind === 'inline-remove' ||
      (edit.kind === 'siblings' &&
        (edit.blocks ?? []).some((binding) => binding.path.length === 1)) ||
      insertedAtomsOfEdit(edit).some((atom) => atom.kind === 'inline-control' || hasImages)
  );
}

/** Capture complete, fully initialized application subtrees after all ordered native stages. */
export function stageTrackedPresence(
  writer: TrackedStageWriter,
  before: OoxmlPackage,
  plan: CanonicalEditPlan,
  attribution: TrackedStageAttribution,
  supplied?: { readonly proposalId: string },
  receipts: readonly ReviewGroupReceipt[] = [],
  images?: InlineImageRuntime
): void {
  const births = plan.operations.flatMap((edit) =>
    insertedAtomsOfEdit(edit).flatMap((atom) =>
      atom.kind === 'inline-control' ? [atom.payload.id] : images ? [images.blockId(atom)] : []
    )
  );
  const removals = plan.operations.filter((edit) => edit.kind === 'inline-remove');
  const roots = plan.operations.flatMap((edit) =>
    edit.kind === 'siblings'
      ? (edit.blocks ?? [])
          .filter((binding) => binding.path.length === 1)
          .map((binding) => binding.payload.id)
      : []
  );
  if (!needsTrackedPresence(plan, !!images)) {
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
    if (edit.kind === 'inline-atoms') continue;
    for (const atom of insertedAtomsOfEdit(edit)) {
      const blockId = atom.kind === 'inline-control' ? atom.payload.id : images?.blockId(atom);
      if (!blockId)
        refuseWrite('The born image has no admitted resource identity.', 'apply-failed');
      const control = index.byId.get(blockId);
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
      let cell: InlinePresenceNative['cell'];
      if (edit.kind !== 'siblings') {
        let owner = part && parent ? parentNodeOf(part, parent.id) : null;
        while (owner && owner.kind !== 'tableCell') owner = parentNodeOf(part!, owner.id);
        if (!owner) refuseWrite('A cell atom has no native owner.', 'apply-failed');
        const cellOwner = owner;
        const row = parentNodeOf(part!, owner.id);
        if (row?.kind !== 'tableRow') refuseWrite('A cell atom has no native row.', 'apply-failed');
        const rowBirth = collectRevisionSites({ ...part!, root: row }).find(
          (site) =>
            site.node.localName === 'ins' &&
            parentNodeOf(part!, site.node.id)?.localName === 'trPr' &&
            parentNodeOf(part!, parentNodeOf(part!, site.node.id)!.id)?.id === row.id
        );
        const cellBirth = collectRevisionSites({ ...part!, root: cellOwner }).find(
          (site) =>
            site.node.localName === 'cellIns' &&
            parentNodeOf(part!, parentNodeOf(part!, site.node.id)?.id ?? '')?.id === cellOwner.id
        );
        const birth = rowBirth ?? cellBirth;
        const topology = birth && nativePropertyAddress(birth.node);
        if (!topology) refuseWrite('The cell atom lost its native topology birth.', 'apply-failed');
        cell = { kind: rowBirth ? 'row' : 'cell', revision: topology };
      }
      bornInline.push({
        blockId,
        partName: control.partName,
        kind: 'ins',
        revision: address,
        gap: matching[0]!.addresses,
        ...(cell ? { cell } : {}),
        ...(edit.kind === 'siblings' && edit.inserted.some((node) => node.type === 'table')
          ? { siblings: true as const }
          : {}),
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
