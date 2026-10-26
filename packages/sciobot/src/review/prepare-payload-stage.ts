import {
  collectRevisionSites,
  contentControlPropertiesOf,
  findNode,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { DEFAULT_PAYLOAD_LIMITS, isBlockId, type BlockPayload } from '../blocks/payload-schema';
import { copyBoundedJson } from '../data/json';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { serializeEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { planParagraphText } from '../diff/paragraph-text';
import { jsonDataKey } from '../diff/target-key';
import { createPendingPropertyProposal } from './property-projection';
import { planDomainProperties } from '../diff/domain-properties';
import { planAtomicBlock } from '../diff/atomic-blocks';
import { readApplicationProposals } from './application-proposals';
import { readPropertyReviewStore, withPropertyReviewRecord } from './property-record-store';
import {
  encodePropertyReviewRecord,
  type PropertyReviewRecord,
  type StoredPropertyBlock,
  type StoredPropertyRevision,
  FIELD_PROPERTY_REVIEW_EXECUTOR,
  ATOMIC_PAYLOAD_CHANGE,
  ATOMIC_REVIEW_EXECUTOR,
} from './property-record-schema';
import { trackedParagraphTextOps, canonicalTrackedDate } from './tracked-inline';
import {
  nativePropertyAddress,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import { payloadFallback, payloadFallbackContains, type PayloadFallback } from './payload-fallback';
import type { PropertyFallbackTarget, TrackPropertiesContext } from './track-properties';

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
export interface PreparedPayloadStage {
  readonly success: true;
  readonly ops: readonly TreeDocOp[];
  readonly complete: (pkg: OoxmlPackage) => {
    readonly pkg: OoxmlPackage;
    readonly receipt: PropertyReviewRecord;
  };
}

/** Validate exact payload/label before-values without a live editor or document mutation. */
export function prepareApplicationPayloadStage(
  pkg: OoxmlPackage,
  current: RevisionToken,
  input: CanonicalEditPlan,
  fallbackTargets: readonly PropertyFallbackTarget[],
  context: TrackPropertiesContext,
  executor: 'properties' | 'atomic'
): PreparedPayloadStage | NativeEditFailure {
  const atomic = executor === 'atomic';
  const copied = copyBoundedJson({ context, fallbackTargets }, DEFAULT_PAYLOAD_LIMITS);
  if (!copied.ok) return fail(copied.reason);
  const data = copied.value as unknown as {
      context: TrackPropertiesContext;
      fallbackTargets: PropertyFallbackTarget[];
    },
    attribution = data.context;
  if (
    !attribution ||
    !isBlockId(attribution.proposalId) ||
    typeof attribution.documentId !== 'string' ||
    !attribution.documentId ||
    typeof attribution.epoch !== 'string' ||
    !attribution.epoch ||
    typeof attribution.author !== 'string' ||
    !attribution.author.trim() ||
    !Array.isArray(data.fallbackTargets)
  )
    return fail('Invalid tracked property context or fallback targets.');
  if (typeof attribution.date !== 'string') return fail('Property proposals need a UTC date.');
  const date = canonicalTrackedDate(attribution.date);
  if (!date) return fail('Property proposals need a valid UTC date.');
  let plan: CanonicalEditPlan;
  try {
    plan = JSON.parse(serializeEditPlan(input)) as CanonicalEditPlan;
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause));
  }
  const pending = createPendingPropertyProposal(attribution.proposalId, plan, attribution.intent);
  if (!pending.success) return pending;
  const part = pkg.parts.get(pkg.mainDocumentPart)!,
    index = buildBlockIndex(pkg),
    payloads = readPayloadStore(pkg),
    records = readPropertyReviewStore(pkg),
    existing = readApplicationProposals(pkg, current);
  if (!existing.success) return existing;
  if (payloads.status !== 'ready' || records.status !== 'ready')
    return fail('Property suggestions need a valid native application store.', 'invalid-document');
  if (records.records.has(attribution.proposalId) || records.removedIds.has(attribution.proposalId))
    return fail('The property proposal ID is already reserved.', 'ambiguous-target');
  const blocks = new Map<string, StoredPropertyBlock>();
  for (const edit of pending.proposal.plan.operations) {
    if (edit.kind !== 'payload' || (atomic ? edit.path.length !== 0 : edit.path.length !== 2))
      return fail('The payload plan does not match this review executor.', 'unsupported-operation');
    const owner = index.byId.get(edit.blockId);
    if (
      !owner ||
      owner.controlId !== edit.controlId ||
      owner.partName !== edit.partName ||
      owner.partName !== part.name
    )
      return fail('A property has no exact current body block owner.', 'invalid-document');
    const definition = blocks.get(owner.id) ?? {
      blockId: owner.id,
      type: owner.payload.type,
      changes: [],
    };
    blocks.set(owner.id, {
      ...definition,
      changes: [
        ...definition.changes,
        {
          name: atomic ? ATOMIC_PAYLOAD_CHANGE : (edit.path[1] as string),
          before: edit.before,
          after: edit.after,
        },
      ],
    });
  }
  // Follow-up ownership of one generated label needs coordinated multi-wave decisions (R14/R19).
  if (
    existing.properties.some((proposal) =>
      proposal.plan.operations.some((edit) => edit.kind === 'payload' && blocks.has(edit.blockId))
    )
  )
    return fail('This block already has a pending property proposal.', 'unsupported-operation');
  if (
    existing.presence.some((record) =>
      record.blocks.some((root) => root.members.some((member) => blocks.has(member.payload.id)))
    )
  )
    return fail('This block belongs to a pending presence proposal.', 'unsupported-operation');
  // Views also support materialized proposals. New writes require exact old-value presence.
  for (const block of blocks.values())
    for (const change of block.changes) {
      const payload = index.byId.get(block.blockId)!.payload,
        attributes = payload.attributes;
      if (
        (atomic ? true : Object.hasOwn(attributes, change.name)) !== change.before.present ||
        (change.before.present &&
          jsonDataKey(atomic ? payload : attributes[change.name]!) !==
            jsonDataKey(change.before.value))
      )
        return fail(
          'Planned property before-values differ from the native baseline.',
          'stale-revision'
        );
    }
  // This validator also runs inside observed native staging. A derived package writer
  // would leak discarded payload projections into the live primitive journal. Validate
  // complete target values through the pure planners; only complete() writes a record.
  for (const block of blocks.values()) {
    if (new Set(block.changes.map((change) => change.name)).size !== block.changes.length)
      return fail('One property proposal addresses a property more than once.', 'ambiguous-target');
    const after = block.changes[0]!.after;
    const checked = atomic
      ? after.present
        ? planAtomicBlock(
            { pkg, revision: current, blockId: block.blockId },
            after.value as unknown as BlockPayload
          )
        : fail('An atomic replacement needs both payloads.')
      : planDomainProperties(
          { pkg, revision: current, blockId: block.blockId },
          {
            set: Object.fromEntries(
              block.changes.flatMap((change) =>
                change.after.present ? [[change.name, change.after.value]] : []
              )
            ),
            remove: block.changes
              .filter((change) => !change.after.present)
              .map((change) => change.name),
          },
          attribution.intent
        );
    if (!checked.success) return checked;
  }
  if (data.fallbackTargets.length !== blocks.size)
    return fail('Each changed block needs one explicit generated fallback target.');
  const ops: TreeDocOp[] = [],
    fallbacks: PayloadFallback[] = [],
    labels: {
      blockId: string;
      before: string;
      after: string;
      textBefore: string;
      textAfter: string;
    }[] = [],
    seen = new Set<string>(),
    beforeSites = collectRevisionSites(part),
    priorKeys = new Set(
      beforeSites.flatMap((site) => {
        const address = nativePropertyAddress(site.node);
        return address ? [propertyAddressKey(part.name, address)] : [];
      })
    );
  // Original offsets remain valid when several inline labels share one paragraph.
  const orderedTargets = [...data.fallbackTargets].sort((a, b) => {
    const left = index.byId.get(a?.blockId),
      right = index.byId.get(b?.blockId);
    const x = left && payloadFallback(part, left.controlId, atomic),
      y = right && payloadFallback(part, right.controlId, atomic);
    return x && y ? x.paragraph.id.localeCompare(y.paragraph.id) || y.offset - x.offset : 0;
  });
  for (const target of orderedTargets) {
    if (
      !target ||
      Object.keys(target).some((key) => !['blockId', 'before', 'after'].includes(key)) ||
      !blocks.has(target.blockId) ||
      seen.has(target.blockId)
    )
      return fail('Fallback targets require unique changed block IDs.');
    seen.add(target.blockId);
    const owner = index.byId.get(target.blockId)!,
      control = findNode(part, owner.controlId),
      fallback = payloadFallback(part, owner.controlId, atomic),
      paragraph = fallback?.paragraph,
      entry = payloads.entries.get(owner.id);
    if (
      !control ||
      control.kind !== 'contentControl' ||
      paragraph?.kind !== 'paragraph' ||
      !fallback ||
      !entry
    )
      return fail(
        'The property block needs one generated native fallback paragraph.',
        'unsupported-operation'
      );
    if (contentControlPropertiesOf(control).dataBinding)
      return fail(
        'Bound property fallbacks need their own native reconciliation.',
        'unsupported-operation'
      );
    const restriction =
      propertyControlRefusal(part, owner.controlId) ?? propertyControlRefusal(part, paragraph.id);
    if (restriction) return fail(restriction, 'unsupported-operation');
    const text = fallback.text;
    if (
      typeof target.before !== 'string' ||
      typeof target.after !== 'string' ||
      text !== target.before
    )
      return fail('The generated fallback before-text no longer matches.', 'stale-revision');
    if (beforeSites.some((site) => payloadFallbackContains(part, fallback, site.node)))
      return fail(
        'The generated fallback already contains pending native revisions.',
        'unsupported-operation'
      );
    const textPlan = planParagraphText(
      {
        revision: current,
        partName: part.name,
        paragraphId: paragraph.id,
        text,
        offset: fallback.offset,
      },
      target.after
    );
    if (!textPlan.success) return textPlan;
    for (const edit of textPlan.plan.operations) {
      if (edit.kind !== 'text') return fail('Unexpected native fallback intention.');
      ops.push(
        ...trackedParagraphTextOps(edit, paragraph, attribution.author, date).map((op) =>
          fallback.inline && op.op === 'insertText' ? { ...op, inside: owner.controlId } : op
        )
      );
    }
    fallbacks.push(fallback);
    labels.push({
      blockId: owner.id,
      before: entry.label,
      after: target.after,
      textBefore: target.before,
      textAfter: target.after,
    });
  }
  const prototype: PropertyReviewRecord = {
    schemaVersion: 1,
    id: attribution.proposalId,
    author: attribution.author,
    date,
    predecessors: [
      ...existing.properties.map((proposal) => proposal.id),
      ...existing.presence.map((record) => record.id),
    ],
    blocks: [...blocks.values()],
    nativeRevisions: [],
    extension: {
      executor: atomic
        ? ATOMIC_REVIEW_EXECUTOR
        : attribution.intent === 'field-edit'
          ? FIELD_PROPERTY_REVIEW_EXECUTOR
          : 'sciobot-properties-v1',
      fallbackLabels: labels,
    },
  };
  try {
    encodePropertyReviewRecord(prototype);
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause));
  }
  return {
    success: true,
    ops,
    complete(working) {
      const changed = working.parts.get(part.name)!,
        revisions = new Map<string, StoredPropertyRevision>();
      for (const site of collectRevisionSites(changed)) {
        if (!fallbacks.some((fallback) => payloadFallbackContains(changed, fallback, site.node)))
          continue;
        const address = nativePropertyAddress(site.node);
        if (
          !address ||
          priorKeys.has(propertyAddressKey(part.name, address)) ||
          address.author !== attribution.author ||
          address.date !== date ||
          site.refused
        )
          throw new Error('Property fallback revision ownership is not exact.');
        revisions.set(propertyAddressKey(part.name, address), {
          partName: part.name,
          ...address,
        });
      }
      if (ops.length && !revisions.size)
        throw new Error('Native property fallback revisions are missing.');
      const receipt = { ...prototype, nativeRevisions: [...revisions.values()] };
      const written = withPropertyReviewRecord(working, receipt);
      if (!written.ok) throw new Error(written.reason);
      return { pkg: written.pkg, receipt };
    },
  };
}
