import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { codeSourceReviewLocks } from './code-source-review';
import { nativeCodePayload } from './code-review-payload';
import type { ReviewGroupPropertyMember } from './groups';
import { ATOMIC_REVIEW_EXECUTOR } from './property-record-schema';
import { readPropertyReviewStore } from './property-record-store';

/** Re-read exact code owners. Serializable ownership receipts never carry unlock authority. */
export function ownedProposalCodeLocks(
  pkg: OoxmlPackage,
  members: readonly ReviewGroupPropertyMember[]
): ReturnType<typeof codeSourceReviewLocks> | null {
  const index = buildBlockIndex(pkg),
    records = readPropertyReviewStore(pkg);
  const ids = new Set<string>();
  for (const member of members) {
    const block = index.byId.get(member.blockId);
    if (!block || !nativeCodePayload(block.payload)) continue;
    const record = records.status === 'ready' && records.records.get(member.proposalId);
    if (
      !record ||
      record.decoded.status !== 'supported' ||
      record.decoded.value.extension?.executor !== ATOMIC_REVIEW_EXECUTOR ||
      !record.decoded.value.blocks.some(
        (saved) => saved.blockId === block.id && saved.type === block.payload.type
      )
    )
      throw Error('The owned code inverse has no exact native proposal.');
    ids.add(block.id);
  }
  return ids.size ? codeSourceReviewLocks(pkg, [...ids]) : null;
}
