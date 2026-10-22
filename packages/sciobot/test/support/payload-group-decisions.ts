import { expect } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { PayloadGroupDecisionRecord } from '../../src/review/payload-group-decision-schema';
import type { PropertyReviewRecord } from '../../src/review/property-record-schema';
import { withPropertyReviewRecord } from '../../src/review/property-record-store';
import {
  readPayloadGroupDecisions,
  payloadGroupDecisionKey,
} from '../../src/review/payload-group-decisions';
import {
  withPayloadGroupDecisionRecord,
  type ReviewGroupStoreWrite,
} from '../../src/review/group-record-store';
import { propertyRecord, propertyRecordFixture } from './property-record';

export const multiBlockProposal: PropertyReviewRecord = {
  ...propertyRecord('shared-proposal'),
  nativeRevisions: [],
  blocks: [
    ...propertyRecord().blocks,
    {
      blockId: 'phase',
      type: 'lesson_phase',
      changes: [
        {
          name: 'phaseType',
          before: { present: true, value: 'evocation' },
          after: { present: true, value: 'reflection' },
        },
      ],
    },
  ],
};
export const decisionReceipt = (
  id = 'rg_alice',
  blockId = 'activity',
  decision: 'accept' | 'reject' = 'accept'
): PayloadGroupDecisionRecord => ({
  schemaVersion: 2,
  id,
  proposalId: multiBlockProposal.id,
  blockId,
  decision,
});
export function decisionFixture() {
  const result = withPropertyReviewRecord(propertyRecordFixture(), multiBlockProposal);
  if (!result.ok) throw new Error(result.reason);
  return result.pkg;
}
export function decisionStates(pkg: OoxmlPackage) {
  const result = readPayloadGroupDecisions(pkg);
  if (result.status !== 'ready') throw new Error(result.reason);
  return [...result.decisions.values()]
    .map((state) => [state.blockId, state.status, [...state.receiptIds]] as const)
    .sort(([left], [right]) => left.localeCompare(right));
}
export function decisionState(pkg: OoxmlPackage, blockId = 'activity') {
  const result = readPayloadGroupDecisions(pkg);
  if (result.status !== 'ready') throw new Error(result.reason);
  return result.decisions.get(payloadGroupDecisionKey(multiBlockProposal.id, blockId));
}
export function mutateDecision(editor: DocxEditorInstance, record: PayloadGroupDecisionRecord) {
  const result = editor.surface!.applyAutomationOps(
    () => [],
    { kind: 'body' },
    [
      (pkg) => {
        const next = withPayloadGroupDecisionRecord(pkg, record);
        if (!next.ok) throw new Error(next.reason);
        return next.pkg;
      },
    ],
    false,
    true
  );
  expect(result.rejected).toBe(false);
  expect(result.committed).toBe(true);
}
export function written(change: ReviewGroupStoreWrite) {
  if (!change.ok) throw new Error(change.reason);
  return change.pkg;
}
