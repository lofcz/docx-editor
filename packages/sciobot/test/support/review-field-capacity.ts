import * as Y from 'yjs';
import {
  REVIEW_FIELD_TARGETS,
  REVIEW_FIELD_INTENTS,
  MAX_REVIEW_FIELD_TARGETS,
  MAX_REVIEW_FIELD_INTENTS,
  fieldAddress,
  encodeFieldRecord,
} from '../../src/collaboration/review-field-codec';

/** Valid inert application rows, independent of the native transport schema. */
export function seedReviewFieldCapacity(
  doc: Y.Doc,
  kind: 'targets' | 'intents',
  remaining: number
) {
  const targets = doc.getMap<string>(REVIEW_FIELD_TARGETS),
    intents = doc.getMap<string>(REVIEW_FIELD_INTENTS);
  doc.transact(() => {
    targets.set('protocol', '1');
    const target = { root: 'sciobot-test-inert-fields', key: 'retired', baseline: null },
      address = fieldAddress(target);
    if (kind === 'targets') {
      const count = MAX_REVIEW_FIELD_TARGETS - remaining - (targets.size - 1);
      for (let i = 0; i < count; i++) {
        const record = { ...target, key: 'retired_' + i };
        targets.set(fieldAddress(record), encodeFieldRecord(record));
      }
    } else {
      targets.set(address, encodeFieldRecord(target));
      const count = MAX_REVIEW_FIELD_INTENTS - remaining - intents._map.size;
      for (let i = 0; i < count; i++)
        intents.set(
          '9000000000:' + i,
          encodeFieldRecord({
            target: address,
            value: null,
            stamp: [0, 9000000000, i],
          })
        );
    }
  }, 'field-capacity-fixture');
}
