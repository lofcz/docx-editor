import * as Y from 'yjs';
import {
  flagAddress,
  MAX_REVIEW_FLAGS,
  MAX_REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_INTENTS,
  REVIEW_FLAG_TARGETS,
} from '../../src/collaboration/review-flag-codec';

/** Valid inert ownership rows outside the native transport schema. Never imported by production. */
export function seedReviewCapacity(doc: Y.Doc, kind: 'targets' | 'intents', remaining: number) {
  const root = doc.getMap<Y.Map<unknown>>('sciobot-test-review-capacity'),
    targets = doc.getMap<string>(REVIEW_FLAG_TARGETS),
    intents = doc.getMap<string>(REVIEW_FLAG_INTENTS);
  if (targets.size > Number(targets.has('protocol')) || intents.size)
    throw new Error('Seed only an unused ownership ledger.');
  doc.transact(() => {
    root.set('retired', new Y.Map());
    const map = root.get('retired')!,
      id = map._item!.id;
    targets.set('protocol', targets.get('protocol') ?? '1');
    if (kind === 'targets') {
      for (let index = 0; index < MAX_REVIEW_FLAGS - remaining; index++) {
        const target = {
          client: id.client,
          clock: id.clock,
          key: `retired_${index}`,
          baseline: 'missing',
        };
        targets.set(flagAddress(target), JSON.stringify(target));
      }
    } else {
      const target = { client: id.client, clock: id.clock, key: 'retired', baseline: 'missing' },
        address = flagAddress(target);
      map.set('retired', true);
      targets.set(address, JSON.stringify(target));
      for (let index = 0; index < MAX_REVIEW_FLAG_INTENTS - remaining; index++)
        intents.set(`1000:${index}`, address);
    }
  }, 'capacity-fixture');
}
