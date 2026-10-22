import { withReviewGroupRevisionReservations } from './group-allocation';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  collectRevisionSites,
  contentControlPropertiesOf,
  findNode,
  paragraphTextOf,
  applyTreeOp,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { DEFAULT_PAYLOAD_LIMITS, isBlockId } from '../blocks/payload-schema';
import { copyBoundedJson } from '../data/json';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { serializeEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { planParagraphText } from '../diff/paragraph-text';
import { jsonDataKey } from '../diff/target-key';
import { readEditorRevision } from '../projection/revision-cache';
import { createPendingPropertyProposal, projectPendingProperties } from './property-projection';
import { readStoredPropertyProposals } from './stored-property-proposals';
import { readPropertyReviewStore, withPropertyReviewRecord } from './property-record-store';
import {
  encodePropertyReviewRecord,
  type PropertyReviewRecord,
  type StoredPropertyBlock,
  type StoredPropertyRevision,
  ATOMIC_PAYLOAD_CHANGE,
  ATOMIC_REVIEW_EXECUTOR,
} from './property-record-schema';
import {
  trackedParagraphTextOps,
  canonicalTrackedDate,
  type TrackTextContext,
} from './tracked-inline';
import {
  nativePropertyAddress,
  paragraphOwner,
  propertyAddressKey,
  propertyFallbackParagraph,
  propertyControlRefusal,
  atomicFallbackParagraph,
} from './property-fallback';
import { persistReviewGroups } from './groups';

/** The application supplies the exact localized label; the adapter never invents content. */
export interface PropertyFallbackTarget {
  readonly blockId: string;
  readonly before: string;
  readonly after: string;
}
export interface TrackPropertiesContext extends TrackTextContext {
  readonly proposalId: string;
}
export type TrackPropertiesResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposal: PropertyReviewRecord | null;
    }
  | NativeEditFailure;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/** Propose attribute deltas and generated Word label edits in one gated native transaction. */
export function trackPropertyPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  fallbackTargets: readonly PropertyFallbackTarget[],
  context: TrackPropertiesContext
): TrackPropertiesResult {
  return trackApplicationPayloadPlan(editor, input, fallbackTargets, context, 'properties');
}

/** Internal shared transaction. The executor selects the stored change semantics. */
export function trackApplicationPayloadPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  fallbackTargets: readonly PropertyFallbackTarget[],
  context: TrackPropertiesContext,
  executor: 'properties' | 'atomic'
): TrackPropertiesResult {
  return withReviewGroupRevisionReservations(editor.surface?.session.currentPackage() ?? null, () =>
    trackApplicationPayloadPlanReserved(editor, input, fallbackTargets, context, executor)
  );
}

function trackApplicationPayloadPlanReserved(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  fallbackTargets: readonly PropertyFallbackTarget[],
  context: TrackPropertiesContext,
  executor: 'properties' | 'atomic'
): TrackPropertiesResult {
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
  const current = readEditorRevision(editor, attribution.documentId, attribution.epoch),
    surface = editor.surface;
  if (!surface || !current) return fail('The native editor is not attached.', 'document-not-ready');
  if (!sameRevision(current, plan.revision))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read the current document before proposing properties.',
      expectedRevision: plan.revision,
      actualRevision: current,
    };
  if (surface.editingMode() === 'view')
    return fail('The document is open for viewing.', 'unsupported-operation');
  if (!plan.operations.length && !data.fallbackTargets.length)
    return { success: true, revision: current, proposal: null };
  const pending = createPendingPropertyProposal(attribution.proposalId, plan);
  if (!pending.success) return pending;
  const pkg = surface.session.currentPackage(),
    part = surface.session.part(),
    index = buildBlockIndex(pkg),
    payloads = readPayloadStore(pkg),
    records = readPropertyReviewStore(pkg),
    existing = readStoredPropertyProposals(pkg, current);
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
    existing.pending.some((proposal) =>
      proposal.plan.operations.some((edit) => edit.kind === 'payload' && blocks.has(edit.blockId))
    )
  )
    return fail('This block already has a pending property proposal.', 'unsupported-operation');
  const checked = projectPendingProperties(pkg, current, [pending.proposal], 'working');
  if (!checked.success) return checked;
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
  if (data.fallbackTargets.length !== blocks.size)
    return fail('Each changed block needs one explicit generated fallback target.');
  const ops: TreeDocOp[] = [],
    paragraphIds = new Set<string>(),
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
  for (const target of data.fallbackTargets) {
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
      paragraph = atomic
        ? atomicFallbackParagraph(part, owner.controlId)
        : propertyFallbackParagraph(part, owner.controlId),
      entry = payloads.entries.get(owner.id);
    if (!control || control.kind !== 'contentControl' || paragraph?.kind !== 'paragraph' || !entry)
      return fail(
        'The property block needs one generated native fallback paragraph.',
        'unsupported-operation'
      );
    if (contentControlPropertiesOf(control).dataBinding)
      return fail(
        'Bound property fallbacks need their own native reconciliation.',
        'unsupported-operation'
      );
    const restriction = propertyControlRefusal(part, paragraph.id);
    if (restriction) return fail(restriction, 'unsupported-operation');
    const text = paragraphTextOf(part, paragraph.id);
    if (
      typeof target.before !== 'string' ||
      typeof target.after !== 'string' ||
      text !== target.before
    )
      return fail('The generated fallback before-text no longer matches.', 'stale-revision');
    if (beforeSites.some((site) => paragraphOwner(part, site.node) === paragraph.id))
      return fail(
        'The generated fallback already contains pending native revisions.',
        'unsupported-operation'
      );
    const textPlan = planParagraphText(
      { revision: current, partName: part.name, paragraphId: paragraph.id, text },
      target.after
    );
    if (!textPlan.success) return textPlan;
    for (const edit of textPlan.plan.operations) {
      if (edit.kind !== 'text') return fail('Unexpected native fallback intention.');
      ops.push(...trackedParagraphTextOps(edit, paragraph, attribution.author, date));
    }
    paragraphIds.add(paragraph.id);
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
    predecessors: existing.pending.map((proposal) => proposal.id),
    blocks: [...blocks.values()],
    nativeRevisions: [],
    extension: {
      executor: atomic ? ATOMIC_REVIEW_EXECUTOR : 'sciobot-properties-v1',
      fallbackLabels: labels,
    },
  };
  try {
    encodePropertyReviewRecord(prototype);
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause));
  }
  let preview = part;
  for (const op of ops) {
    const result = applyTreeOp(preview, op);
    if (!result.ok)
      return fail(`Native property fallback refused: ${result.reason}.`, 'unsupported-operation');
    preview = result.part;
  }
  let receipt: PropertyReviewRecord | null = null;
  try {
    const result = surface.applyAutomationOps(
      () => ops,
      { kind: 'body' },
      [
        (working) => {
          const changed = working.parts.get(part.name)!,
            revisions = new Map<string, StoredPropertyRevision>();
          for (const site of collectRevisionSites(changed)) {
            if (!paragraphIds.has(paragraphOwner(changed, site.node) ?? '')) continue;
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
          receipt = { ...prototype, nativeRevisions: [...revisions.values()] };
          const written = withPropertyReviewRecord(working, receipt);
          if (!written.ok) throw new Error(written.reason);
          return persistReviewGroups(
            pkg,
            written.pkg,
            [],
            editor.snapshot().collaborationStatus === 'inactive'
          );
        },
      ],
      true,
      true,
      'after'
    );
    if (result.rejected || !result.committed)
      return fail(String(result.reason ?? 'The property proposal did not commit.'), 'apply-failed');
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'apply-failed');
  }
  return {
    success: true,
    revision: { ...current, revision: surface.session.packageRevision() },
    proposal: receipt,
  };
}
