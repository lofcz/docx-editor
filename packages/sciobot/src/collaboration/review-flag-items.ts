// Pinned MIT Yjs item identities. No provider schema, field name, or implementation is read.
import * as Y from 'yjs';
import { ReviewFlagError, type ReviewFlagTarget } from './review-flag-codec';

export function flagMap(doc: Y.Doc, target: ReviewFlagTarget): Y.Map<unknown> | null {
  let id = Y.createID(target.client, target.clock);
  const visited = new Set<string>();
  for (let depth = 0; depth < 256; depth++) {
    const key = `${id.client}:${id.clock}`;
    if (visited.has(key)) throw new ReviewFlagError('Review flag target lineage contains a cycle.');
    visited.add(key);
    const structs = doc.store.clients.get(id.client);
    if (!structs?.length || Y.getState(doc.store, id.client) <= id.clock) return null;
    const item = Y.getItem(doc.store, id);
    if (!(item instanceof Y.Item)) return null;
    if (item.redone) {
      id = item.redone;
      continue;
    }
    if (item.deleted) return null;
    if (!(item.content instanceof Y.ContentType) || !(item.content.type instanceof Y.Map))
      throw new ReviewFlagError('Review flag target must name a shared map.');
    return item.content.type;
  }
  throw new ReviewFlagError('Review flag target lineage exceeds its limit.');
}
