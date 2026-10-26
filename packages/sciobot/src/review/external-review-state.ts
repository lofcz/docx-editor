import { payloadFallback, payloadFallbackContains, type PayloadFallback } from './payload-fallback';
import { FIELD_PROPERTY_REVIEW_EXECUTOR } from './property-record-schema';
import {
  applyTreeOp,
  collectRevisionSites,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import type { NativeEditFailure } from '../contracts/document';
import type { JsonObject } from '../data/json';
import {
  decodeBlockPayload,
  encodeBlockPayload,
  type BlockPayload,
} from '../blocks/payload-schema';
import { domainDiffStrategy, sameContainerIdentity } from '../diff/domain-strategies';
import { DOMAIN_BLOCK_TAGS } from '../projection/block-tags';
import { MDX_BLOCK, MDX_INLINE, serializeOpaqueMdx } from '../projection/mdx';
import { jsonDataKey } from '../diff/target-key';
import { readPropertyReviewStore } from './property-record-store';
import {
  ATOMIC_PAYLOAD_CHANGE,
  ATOMIC_REVIEW_EXECUTOR,
  type PropertyReviewRecord,
  type StoredPropertyBlock,
} from './property-record-schema';
import { readPayloadGroupDecisions, payloadGroupDecisionKey } from './payload-group-decisions';
import {
  nativePropertyAddress,
  paragraphOwner,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';

export type ExternalReviewStatus = 'pending' | 'accepted' | 'rejected' | 'resolved' | 'conflict';
export type ExternalReviewConflict =
  | 'conflicting-decisions'
  | 'overlapping-proposals'
  | 'invalid-execution-metadata'
  | 'missing-owner'
  | 'restricted-owner'
  | 'changed-committed-value'
  | 'foreign-revision'
  | 'moved-revision'
  | 'changed-fallback';
export interface ExternalReviewEntry {
  readonly proposalId: string;
  readonly blockId: string;
  readonly status: ExternalReviewStatus;
  readonly conflict: ExternalReviewConflict | null;
}
export interface ExternalReviewReport {
  readonly entries: readonly ExternalReviewEntry[];
  readonly hasConflicts: boolean;
}
export type ExternalReviewRead =
  | { readonly success: true; readonly report: ExternalReviewReport }
  | NativeEditFailure;
export interface ExternalFallbackLabel {
  readonly before: string;
  readonly after: string;
  readonly textBefore: string;
  readonly textAfter: string;
}
export interface ExternalReviewTarget {
  readonly record: PropertyReviewRecord;
  readonly block: StoredPropertyBlock;
  readonly label: ExternalFallbackLabel;
  readonly atomic: boolean;
  readonly decision: 'accept' | 'reject';
}
interface Candidate {
  readonly record: PropertyReviewRecord;
  readonly block: StoredPropertyBlock;
  readonly atomic: boolean;
  readonly label: ExternalFallbackLabel | null;
  readonly part: OoxmlPart | null;
  readonly paragraphId: string | null;
  readonly fallback: PayloadFallback | null;
  readonly controlId: string | null;
  readonly keys: ReadonlySet<string>;
  conflict: ExternalReviewConflict | null;
  native: readonly LocatedSite[];
}
interface LocatedSite {
  readonly site: ReturnType<typeof collectRevisionSites>[number];
  readonly key: string | null;
  readonly paragraph: string | null;
  readonly location: string;
}
const locationKey = (part: string, paragraph: string | null) => JSON.stringify([part, paragraph]);
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-document'
): NativeEditFailure => ({ success: false, code, error });

function labels(record: PropertyReviewRecord) {
  const result = new Map<string, ExternalFallbackLabel>();
  const raw = record.extension?.fallbackLabels;
  if (!Array.isArray(raw) || raw.length !== record.blocks.length) return null;
  for (const entry of raw) {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) return null;
    const value = entry as JsonObject;
    if (
      Object.keys(value).some(
        (key) => !['blockId', 'before', 'after', 'textBefore', 'textAfter'].includes(key)
      ) ||
      typeof value.blockId !== 'string' ||
      !record.blocks.some((block) => block.blockId === value.blockId) ||
      typeof value.before !== 'string' ||
      typeof value.after !== 'string' ||
      typeof value.textBefore !== 'string' ||
      typeof value.textAfter !== 'string' ||
      value.after !== value.textAfter ||
      result.has(value.blockId)
    )
      return null;
    result.set(value.blockId, {
      before: value.before,
      after: value.after,
      textBefore: value.textBefore,
      textAfter: value.textAfter,
    });
  }
  return result;
}

/** Match the existing domain/atomic planners without constructing another document model. */
function validPayloadChange(
  payload: BlockPayload,
  block: StoredPropertyBlock,
  atomic: boolean
): boolean {
  try {
    if (atomic) {
      const pair = block.changes[0];
      if (!pair || !pair.before.present || !pair.after.present) return false;
      const before = decodeBlockPayload(JSON.stringify(pair.before.value)),
        after = decodeBlockPayload(JSON.stringify(pair.after.value));
      if (
        before.status !== 'supported' ||
        after.status !== 'supported' ||
        before.value.id !== block.blockId ||
        after.value.id !== block.blockId ||
        before.value.type !== block.type ||
        after.value.type !== block.type
      )
        return false;
      if (block.type === MDX_BLOCK || block.type === MDX_INLINE) {
        serializeOpaqueMdx(before.value);
        serializeOpaqueMdx(after.value);
        return true;
      }
      return domainDiffStrategy({ type: block.type })?.kind === 'atomic';
    }
    const definition = DOMAIN_BLOCK_TAGS.get(block.type),
      strategy = domainDiffStrategy({ type: block.type });
    if (definition?.kind !== 'container' || strategy?.kind !== 'container') return false;
    const names = new Set(block.changes.map((change) => change.name)),
      probe = Object.fromEntries(
        [...names].map((name) => [
          name,
          name === 'bookId' ? 1 : name === 'trust' ? 'verified' : 'probe',
        ])
      ),
      supported = new Set(
        definition.attributes({ ...payload, attributes: probe }).map((binding) => binding.property)
      );
    if ([...names].some((name) => !supported.has(name))) return false;
    const after = Object.fromEntries([
      ...Object.entries(payload.attributes).filter(([name]) => !names.has(name)),
      ...block.changes.flatMap((change) =>
        change.after.present ? [[change.name, change.after.value]] : []
      ),
    ]);
    if (
      !sameContainerIdentity(
        { ...payload.attributes, type: block.type },
        { ...after, type: block.type }
      )
    )
      return false;
    encodeBlockPayload({ ...payload, attributes: after });
    return true;
  } catch {
    return false;
  }
}

/** Pure canonical inspection. The returned targets are internal, never stored or replayed. */
export function prepareExternalReview(pkg: OoxmlPackage):
  | {
      readonly success: true;
      readonly report: ExternalReviewReport;
      readonly targets: readonly ExternalReviewTarget[];
    }
  | NativeEditFailure {
  const store = readPropertyReviewStore(pkg);
  if (store.status === 'invalid') return fail(store.reason);
  const dispositions = readPayloadGroupDecisions(pkg);
  if (dispositions.status === 'invalid') return fail(dispositions.reason);
  const entries: ExternalReviewEntry[] = [],
    targets: ExternalReviewTarget[] = [];
  if (store.status === 'absent')
    return {
      success: true,
      report: Object.freeze({ entries: Object.freeze(entries), hasConflicts: false }),
      targets: Object.freeze(targets),
    };
  const index = buildBlockIndex(pkg),
    payloads = readPayloadStore(pkg);
  if (payloads.status !== 'ready')
    return fail('External review needs a valid application payload store.');
  const records: PropertyReviewRecord[] = [];
  for (const entry of store.records.values()) {
    if (entry.decoded.status !== 'supported')
      return fail(
        'Unsupported property proposal version in external review.',
        'unsupported-operation'
      );
    records.push(entry.decoded.value);
  }
  const remaining = new Map(records.map((record) => [record.id, record]));
  while (remaining.size) {
    const ready = [...remaining.values()].filter((record) =>
      record.predecessors.every((id) => !remaining.has(id))
    );
    if (!ready.length) return fail('Property proposal predecessors contain a cycle.');
    for (const record of ready) remaining.delete(record.id);
  }
  const candidates: Candidate[] = [],
    counts = new Map<string, number>();
  for (const record of records) {
    const atomic = record.extension?.executor === ATOMIC_REVIEW_EXECUTOR,
      execution =
        atomic ||
        record.extension?.executor === 'sciobot-properties-v1' ||
        record.extension?.executor === FIELD_PROPERTY_REVIEW_EXECUTOR,
      fallback = labels(record),
      keys = new Set(
        record.nativeRevisions.map((revision) => propertyAddressKey(revision.partName, revision))
      ),
      ownership = record.nativeRevisions.every(
        (revision) => revision.author === record.author && revision.date === record.date
      );
    for (const block of record.blocks) {
      const disposition = dispositions.decisions.get(
        payloadGroupDecisionKey(record.id, block.blockId)
      );
      if (disposition && disposition.status !== 'conflict') {
        entries.push(
          Object.freeze({
            proposalId: record.id,
            blockId: block.blockId,
            status: 'resolved',
            conflict: null,
          })
        );
        continue;
      }
      counts.set(block.blockId, (counts.get(block.blockId) ?? 0) + 1);
      const owner = index.byId.get(block.blockId),
        part = owner ? pkg.parts.get(owner.partName) : undefined,
        nativeFallback = owner && part ? payloadFallback(part, owner.controlId, atomic) : null,
        paragraph = nativeFallback?.paragraph,
        payload = payloads.entries.get(block.blockId),
        label = fallback?.get(block.blockId) ?? null;
      let conflict: ExternalReviewConflict | null = disposition ? 'conflicting-decisions' : null;
      if (
        !execution ||
        !fallback ||
        !label ||
        !ownership ||
        (label.textBefore !== label.textAfter && record.nativeRevisions.length === 0) ||
        (atomic && (block.changes.length !== 1 || block.changes[0]!.name !== ATOMIC_PAYLOAD_CHANGE))
      )
        conflict ??= 'invalid-execution-metadata';
      if (!owner || owner.payload.type !== block.type || !part || !paragraph || !payload)
        conflict ??= 'missing-owner';
      if (
        part &&
        paragraph &&
        (propertyControlRefusal(part, owner!.controlId) ??
          propertyControlRefusal(part, paragraph.id))
      )
        conflict ??= 'restricted-owner';
      if (owner && payload && label) {
        if (payload.label !== label.before) conflict ??= 'changed-committed-value';
        for (const change of block.changes) {
          const value = atomic ? owner.payload : owner.payload.attributes[change.name],
            present = atomic || Object.hasOwn(owner.payload.attributes, change.name);
          if (
            present !== change.before.present ||
            (change.before.present && jsonDataKey(value!) !== jsonDataKey(change.before.value))
          )
            conflict ??= 'changed-committed-value';
        }
      }
      candidates.push({
        record,
        block,
        atomic,
        label,
        part: part ?? null,
        paragraphId: paragraph?.id ?? null,
        fallback: nativeFallback,
        controlId: owner?.controlId ?? null,
        keys,
        conflict,
        native: [],
      });
    }
  }
  const byAddress = new Map<string, LocatedSite[]>(),
    byParagraph = new Map<string, LocatedSite[]>();
  for (const part of pkg.parts.values())
    for (const site of collectRevisionSites(part)) {
      const address = nativePropertyAddress(site.node),
        paragraph = paragraphOwner(part, site.node),
        key = address ? propertyAddressKey(part.name, address) : null,
        location = locationKey(part.name, paragraph),
        located = { site, key, paragraph, location };
      const local = byParagraph.get(location) ?? [];
      local.push(located);
      byParagraph.set(location, local);
      if (key) {
        const addressed = byAddress.get(key) ?? [];
        addressed.push(located);
        byAddress.set(key, addressed);
      }
    }
  const moved = new Set<string>();
  for (const record of records)
    for (const revision of record.nativeRevisions) {
      for (const site of byAddress.get(propertyAddressKey(revision.partName, revision)) ?? [])
        if (
          !candidates.some(
            (candidate) =>
              candidate.record.id === record.id &&
              candidate.part &&
              candidate.fallback &&
              candidate.part.name === revision.partName &&
              payloadFallbackContains(candidate.part, candidate.fallback, site.site.node)
          )
        )
          moved.add(record.id);
    }
  const selectedByPart = new Map<string, Set<string>>();
  for (const candidate of candidates) {
    if ((counts.get(candidate.block.blockId) ?? 0) > 1)
      candidate.conflict ??= 'overlapping-proposals';
    if (moved.has(candidate.record.id)) candidate.conflict ??= 'moved-revision';
    const payload = index.byId.get(candidate.block.blockId)?.payload;
    if (
      !candidate.conflict &&
      payload &&
      !validPayloadChange(payload, candidate.block, candidate.atomic)
    )
      candidate.conflict = 'invalid-execution-metadata';
    if (!candidate.part || !candidate.paragraphId) continue;
    const local = (
        byParagraph.get(locationKey(candidate.part.name, candidate.paragraphId)) ?? []
      ).filter(
        ({ site }) =>
          candidate.fallback &&
          payloadFallbackContains(candidate.part!, candidate.fallback, site.node)
      ),
      owned = local.filter((site) => site.key && candidate.keys.has(site.key));
    candidate.native = owned;
    if (
      owned.length !== local.length ||
      local.some(({ site }) => site.refused || !['ins', 'del'].includes(site.node.localName))
    )
      candidate.conflict ??= 'foreign-revision';
    if (candidate.conflict || !owned.length) continue;
    const selected = selectedByPart.get(candidate.part.name) ?? new Set<string>();
    for (const { site } of owned) selected.add(site.node.id);
    selectedByPart.set(candidate.part.name, selected);
  }
  // These temporary native decisions classify incomplete review. Their parts are never installed.
  const views = new Map<string, { original: OoxmlPart | null; proposed: OoxmlPart | null }>();
  for (const [name, selected] of selectedByPart) {
    const part = pkg.parts.get(name)!,
      siteNodeIds = [...selected],
      original = applyTreeOp(part, { op: 'rejectAllRevisions', siteNodeIds }),
      proposed = applyTreeOp(part, { op: 'acceptAllRevisions', siteNodeIds });
    views.set(name, {
      original: original.ok ? original.part : null,
      proposed: proposed.ok ? proposed.part : null,
    });
  }
  for (const candidate of candidates) {
    const { record, block, part, paragraphId, label } = candidate;
    let status: ExternalReviewStatus = 'conflict';
    if (!candidate.conflict && part && paragraphId && label) {
      if (candidate.native.length) {
        const view = views.get(part.name);
        if (
          view?.original &&
          view.proposed &&
          payloadFallback(view.original, candidate.controlId!, candidate.atomic)?.text ===
            label.textBefore &&
          payloadFallback(view.proposed, candidate.controlId!, candidate.atomic)?.text ===
            label.textAfter
        )
          status = 'pending';
        else candidate.conflict = 'changed-fallback';
      } else {
        const text = payloadFallback(part, candidate.controlId!, candidate.atomic)?.text;
        if (label.textBefore === label.textAfter && text === label.textBefore) status = 'pending';
        else if (text === label.textAfter) status = 'accepted';
        else if (text === label.textBefore) status = 'rejected';
        else candidate.conflict = 'changed-fallback';
      }
    }
    if (candidate.conflict) status = 'conflict';
    entries.push(
      Object.freeze({
        proposalId: record.id,
        blockId: block.blockId,
        status,
        conflict: candidate.conflict,
      })
    );
    if ((status === 'accepted' || status === 'rejected') && label)
      targets.push(
        Object.freeze({
          record,
          block,
          label,
          atomic: candidate.atomic,
          decision: status === 'accepted' ? 'accept' : 'reject',
        })
      );
  }
  entries.sort(
    (a, b) => a.proposalId.localeCompare(b.proposalId) || a.blockId.localeCompare(b.blockId)
  );
  return {
    success: true,
    report: Object.freeze({
      entries: Object.freeze(entries),
      hasConflicts: entries.some((entry) => entry.status === 'conflict'),
    }),
    targets: Object.freeze(targets),
  };
}

/** No group state is hidden or modified when external content is ambiguous. */
export function inspectExternalReview(pkg: OoxmlPackage): ExternalReviewRead {
  const prepared = prepareExternalReview(pkg);
  return prepared.success ? Object.freeze({ success: true, report: prepared.report }) : prepared;
}
