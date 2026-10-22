import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { readReviewGroupStore } from './group-record-store';

export interface PayloadGroupDecisionState {
  readonly proposalId: string;
  readonly blockId: string;
  readonly status: 'accepted' | 'rejected' | 'conflict';
  readonly receiptIds: readonly string[];
}
export type PayloadGroupDecisionsRead =
  | { readonly status: 'invalid'; readonly reason: string }
  | {
      readonly status: 'ready';
      /** Absence means pending. Concurrent opposite dispositions require explicit reconciliation. */
      readonly decisions: ReadonlyMap<string, PayloadGroupDecisionState>;
    };
export const payloadGroupDecisionKey = (proposalId: string, blockId: string) =>
  JSON.stringify([proposalId, blockId]);

/** Pure reads preserve every concurrent receipt. No actor or timestamp selects a winner. */
export function readPayloadGroupDecisions(pkg: OoxmlPackage): PayloadGroupDecisionsRead {
  const source = readReviewGroupStore(pkg);
  if (source.status === 'invalid') return source;
  const decisions = new Map<string, PayloadGroupDecisionState>();
  if (source.status === 'absent') return { status: 'ready', decisions };
  for (const entry of source.records.values()) {
    if (entry.decoded.status === 'unsupported-version')
      return { status: 'invalid', reason: 'unsupported-review-group-version' };
    if (entry.decoded.status !== 'payload-decision') continue;
    const receipt = entry.decoded.value,
      key = payloadGroupDecisionKey(receipt.proposalId, receipt.blockId),
      previous = decisions.get(key),
      status = receipt.decision === 'accept' ? 'accepted' : 'rejected';
    decisions.set(
      key,
      Object.freeze({
        proposalId: receipt.proposalId,
        blockId: receipt.blockId,
        status: previous && previous.status !== status ? 'conflict' : status,
        receiptIds: Object.freeze([...(previous?.receiptIds ?? []), receipt.id].sort()),
      })
    );
  }
  return { status: 'ready', decisions };
}
