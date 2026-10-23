import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken, DocumentView } from '../contracts/document';
import {
  BLOCK_PRESENCE_EXECUTOR,
  decodeBlockPresenceRecord,
  type BlockPresenceRecord,
} from './block-presence-schema';
import { readPropertyReviewStore, type StoredPropertyRecord } from './property-record-store';
import { bindStoredPropertyProposals } from './stored-property-proposals';
import { readPayloadGroupDecisions, payloadGroupDecisionKey } from './payload-group-decisions';
import { bindBlockPresenceRecord } from './block-presence-plan';
import { projectBlockPresence } from './block-presence-projection';
import type { PendingPropertyProposal } from './property-projection';

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-document'
): NativeEditFailure => ({ success: false, error, code });
export interface ApplicationProposals {
  readonly success: true;
  readonly properties: readonly PendingPropertyProposal[];
  readonly presence: readonly BlockPresenceRecord[];
}
/** Admit both formats without interpreting a presence proposal as attribute changes. */
export function readApplicationProposals(
  pkg: OoxmlPackage,
  revision: RevisionToken
): ApplicationProposals | NativeEditFailure {
  const store = readPropertyReviewStore(pkg);
  if (store.status === 'invalid') return fail(store.reason);
  if (store.status === 'absent') return { success: true, properties: [], presence: [] };
  if ([...store.retainedRecords.values()].every((entry) => entry.decoded.status === 'supported')) {
    const legacy = bindStoredPropertyProposals(pkg, revision, store.records);
    return legacy.success ? { success: true, properties: legacy.pending, presence: [] } : legacy;
  }
  const records = new Map<string, { readonly predecessors: readonly string[] }>();
  const presence = new Map<string, BlockPresenceRecord>();
  for (const entry of store.retainedRecords.values()) {
    if (entry.decoded.status === 'supported') records.set(entry.id, entry.decoded.value);
    else {
      const decoded = decodeBlockPresenceRecord(entry.raw);
      if (decoded.status !== 'supported' || decoded.value.executor !== BLOCK_PRESENCE_EXECUTOR)
        return fail('Unsupported application proposal version.', 'unsupported-operation');
      records.set(entry.id, decoded.value);
      if (store.records.has(entry.id)) presence.set(entry.id, decoded.value);
    }
  }
  const remaining = new Map(records),
    ancestry = new Map<string, ReadonlySet<string>>();
  while (remaining.size) {
    const ready = [...remaining].filter(([, record]) =>
      record.predecessors.every((id) => !remaining.has(id))
    );
    if (!ready.length) return fail('Application proposal predecessors contain a cycle.');
    for (const [id, record] of ready) {
      const ancestors = new Set(record.predecessors);
      for (const parent of record.predecessors)
        for (const ancestor of ancestry.get(parent) ?? []) ancestors.add(ancestor);
      ancestry.set(id, ancestors);
      remaining.delete(id);
    }
  }
  const legacy = new Map<string, StoredPropertyRecord>();
  for (const entry of store.records.values())
    if (entry.decoded.status === 'supported')
      legacy.set(entry.id, {
        ...entry,
        decoded: {
          status: 'supported',
          value: { ...entry.decoded.value, predecessors: [...ancestry.get(entry.id)!].sort() },
        },
      });
  const boundProperties = bindStoredPropertyProposals(pkg, revision, legacy);
  if (!boundProperties.success) return boundProperties;
  const dispositions = readPayloadGroupDecisions(pkg);
  if (dispositions.status === 'invalid') return fail(dispositions.reason);
  const selected: BlockPresenceRecord[] = [],
    owners = new Set(
      boundProperties.pending.flatMap((proposal) =>
        proposal.plan.operations.flatMap((edit) => (edit.kind === 'payload' ? [edit.blockId] : []))
      )
    );
  for (const record of presence.values()) {
    const blocks = [];
    for (const block of record.blocks) {
      const decision = dispositions.decisions.get(
        payloadGroupDecisionKey(record.id, block.blockId)
      );
      if (decision?.status === 'conflict')
        return fail('Conflicting presence decisions need reconciliation.', 'ambiguous-target');
      if (decision) continue;
      for (const member of block.members) {
        if (owners.has(member.payload.id))
          return fail(
            'Pending application proposals have overlapping presence ownership.',
            'ambiguous-target'
          );
        owners.add(member.payload.id);
      }
      blocks.push(block);
    }
    if (!blocks.length) continue;
    const pending = { ...record, blocks };
    const bound = bindBlockPresenceRecord(pkg, pending);
    if (!bound.success) return bound;
    selected.push(bound.record);
  }
  return { success: true, properties: boundProperties.pending, presence: Object.freeze(selected) };
}
/** Validate all pending owners before deriving any absent-side view. */
export function projectApplicationPresence(
  pkg: OoxmlPackage,
  pending: readonly BlockPresenceRecord[],
  kind: DocumentView
): { readonly success: true; readonly pkg: OoxmlPackage } | NativeEditFailure {
  for (const record of pending) {
    const bound = bindBlockPresenceRecord(pkg, record);
    if (!bound.success) return bound;
  }
  let projected = pkg;
  for (const record of pending) {
    const result = projectBlockPresence(projected, record, kind);
    if (!result.success) return result;
    projected = result.pkg;
  }
  return { success: true, pkg: projected };
}
