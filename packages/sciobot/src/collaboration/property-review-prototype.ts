import * as Y from 'yjs';
import type { attachOwnedReviewFields, ReviewFieldReceipt } from './review-fields';
import type { OwnedReviewFlags } from './review-flags';
import type { ReviewFlagReceipt } from './review-flag-codec';

/** Internal integration probe. Group binding and admission are required before production use. */
export interface PropertyReviewReceipt {
  readonly fields: readonly ReviewFieldReceipt[];
  readonly flags: readonly ReviewFlagReceipt[];
}
export const PROPERTY_REVIEW_PROTOTYPE = Symbol.for('sciobot.docx.property-review.prototype.1');
export interface PropertyReviewPrototype {
  latest(): PropertyReviewReceipt | null;
  withdrawDuringPublication<T>(receipt: PropertyReviewReceipt, run: () => T): T;
}
export function createPropertyReviewPrototype(
  doc: Y.Doc,
  fields: ReturnType<typeof attachOwnedReviewFields>,
  flags: OwnedReviewFlags
) {
  let latest: PropertyReviewReceipt | null = null;
  let withdrawing = false;
  return Object.freeze({
    note(receipt: PropertyReviewReceipt) {
      latest = Object.freeze(receipt);
    },
    latest: () => latest,
    withdrawDuringPublication<T>(receipt: PropertyReviewReceipt, run: () => T): T {
      if (withdrawing || doc._transaction || doc.store.pendingStructs || doc.store.pendingDs)
        throw new Error('Property withdrawal needs an idle complete native publication scope.');
      // Validate both ownership kinds before the callback can commit canonical state.
      fields.assertRetraction(receipt.fields);
      flags.assertRetraction(receipt.flags);
      let applied = false;
      const before = (transaction: Y.Transaction) => {
        const origin = transaction.origin;
        if (
          applied ||
          !transaction.local ||
          !origin ||
          typeof origin !== 'object' ||
          origin.kind !== 'docx-document-local'
        )
          return;
        applied = true;
        flags.retractInTransaction(receipt.flags, () =>
          fields.retractInTransaction(receipt.fields, () => {})
        );
      };
      withdrawing = true;
      doc.on('beforeTransaction', before);
      try {
        const value = run();
        if (!applied) throw new Error('Property withdrawal did not publish a native transaction.');
        return value;
      } finally {
        doc.off('beforeTransaction', before);
        withdrawing = false;
      }
    },
  });
}
