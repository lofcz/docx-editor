// A fresh process verifies both global interceptor installation orders.
import assert from 'node:assert/strict';
import * as Y from 'yjs';
import {
  attachOwnedReviewFields,
  type ReviewFieldReceipt,
} from '../../src/collaboration/review-fields';
import { attachOwnedReviewFlags } from '../../src/collaboration/review-flags';
import type { ReviewFlagReceipt } from '../../src/collaboration/review-flag-codec';
const doc = new Y.Doc();
const order = process.argv[2];
let fields: ReturnType<typeof attachOwnedReviewFields>;
let flags: ReturnType<typeof attachOwnedReviewFlags>;
if (order === 'fields-first') {
  fields = attachOwnedReviewFields(doc, { protocol: '2' });
  flags = attachOwnedReviewFlags(doc, { protocol: '3' });
} else {
  assert.equal(order, 'flags-first');
  flags = attachOwnedReviewFlags(doc, { protocol: '3' });
  fields = attachOwnedReviewFields(doc, { protocol: '2' });
}
const root = doc.getMap<unknown>('native'),
  node = new Y.Map<unknown>();
root.set('shape', 'original');
root.set('node', node);
fields.enable();
const history = new Y.UndoManager(root, { captureTimeout: 0 });
let fieldReceipt: readonly ReviewFieldReceipt[] = [],
  flagReceipt: readonly ReviewFlagReceipt[] = [];
try {
  flags.capture(
    () =>
      fields.capture(
        () =>
          doc.transact(() => {
            root.set('shape', 'proposal');
            node.set('removed', true);
          }),
        1,
        (items) => {
          fieldReceipt = items;
        }
      ),
    1,
    (items) => {
      flagReceipt = items;
    }
  );
  assert.equal(fieldReceipt.length, 1);
  assert.equal(flagReceipt.length, 1);
  doc.transact(() =>
    flags.retractInTransaction(flagReceipt, () =>
      fields.retractInTransaction(
        fieldReceipt,
        () => {
          root.set('shape', 'original');
          root.set('decision', 'withdrawn');
        },
        1
      )
    )
  );
  assert.equal(root.get('shape'), 'original');
  assert.equal(node.has('removed'), false);
  assert.notEqual(history.undo(), null);
  assert.equal(root.get('shape'), 'proposal');
  assert.equal(node.get('removed'), true);
  assert.notEqual(history.redo(), null);
  assert.equal(root.get('shape'), 'original');
  assert.equal(node.has('removed'), false);
} finally {
  history.destroy();
  fields.dispose();
  flags.dispose();
  doc.destroy();
}
