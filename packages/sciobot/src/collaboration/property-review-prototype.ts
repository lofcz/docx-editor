import { createPropertyGroupBindings } from './property-group-bindings';
import {
  PROPERTY_REVIEW_PROTOTYPE,
  type PropertyGroupPublication,
} from '../review/property-history-contract';
export { PROPERTY_REVIEW_PROTOTYPE };
import * as Y from 'yjs';
import type { attachOwnedReviewFields, ReviewFieldReceipt } from './review-fields';
import type { OwnedReviewFlags } from './review-flags';
import type { ReviewFlagReceipt } from './review-flag-codec';

/** Internal integration probe. Group binding and admission are required before production use. */
export interface PropertyReviewReceipt {
  readonly fields: readonly ReviewFieldReceipt[];
  readonly flags: readonly ReviewFlagReceipt[];
}
export interface PropertyReviewPrototype {
  latest(): PropertyReviewReceipt | null;
  read(groupId: string, membersKey: string): PropertyReviewReceipt | null;
  withdrawDuringPublication<T>(
    receipt: PropertyReviewReceipt,
    run: () => T,
    flagWrites?: number,
    fieldWrites?: number
  ): T;
}
export function createPropertyReviewPrototype(
  doc: Y.Doc,
  fields: ReturnType<typeof attachOwnedReviewFields>,
  flags: OwnedReviewFlags
) {
  const groups = createPropertyGroupBindings(doc);
  let latest: PropertyReviewReceipt | null = null;
  let withdrawing = false;
  return Object.freeze({
    assertHealthy: groups.assertHealthy,
    dispose: groups.dispose,
    read: groups.read,
    captureGroups<T>(
      run: () => T,
      publication: PropertyGroupPublication,
      fieldWrites: number,
      flagWrites: number
    ): T {
      groups.reserve(publication.preview, fieldWrites, flagWrites);
      const transact = doc.transact.bind(doc);
      let published = false;
      // Wrap the local publication callback itself. beforeObserverCalls runs after
      // Yjs clears _transaction and would create a second, non-atomic publication.
      const wrapped: typeof doc.transact = (runTransaction, origin) => {
        const outermost = doc._transaction === null;
        return transact((transaction) => {
          const value = runTransaction(transaction);
          if (
            outermost &&
            !published &&
            transaction.local &&
            transaction.origin?.kind === 'docx-document-local'
          ) {
            published = true;
            groups.publish(publication.current(), {
              fields: fields.currentReceipt(),
              flags: flags.currentReceipt(),
            });
          }
          return value;
        }, origin);
      };
      const previous = doc.transact;
      doc.transact = wrapped;
      try {
        return run();
      } finally {
        doc.transact = previous;
      }
    },
    note(receipt: PropertyReviewReceipt) {
      latest = Object.freeze(receipt);
    },
    latest: () => latest,
    withdrawDuringPublication<T>(
      receipt: PropertyReviewReceipt,
      run: () => T,
      flagWrites = 0,
      fieldWrites = 0
    ): T {
      if (withdrawing || doc._transaction || doc.store.pendingStructs || doc.store.pendingDs)
        throw new Error('Property withdrawal needs an idle complete native publication scope.');
      // Validate both ownership kinds before the callback can commit canonical state.
      fields.assertRetraction(receipt.fields, fieldWrites);
      flags.assertRetraction(receipt.flags, flagWrites);
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
