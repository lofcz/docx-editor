import * as Y from 'yjs';
import { decodeLineage, OWNED_TEXT_LINEAGE, OwnedTextError, type IdRange } from './ownership-codec';

export interface CharacterSpan {
  readonly index: number;
  readonly length: number;
  readonly client: number;
  readonly clock: number;
}
export function visibleSpans(text: Y.Text): CharacterSpan[] {
  const output: CharacterSpan[] = [];
  let index = 0;
  for (let item = text._start; item; item = item.right) {
    if (item.deleted || !item.countable) continue;
    output.push({ index, length: item.length, client: item.id.client, clock: item.id.clock });
    index += item.length;
  }
  return output;
}
export function observedRanges(text: Y.Text, index: number, length: number): IdRange[] {
  const output: [number, number, number][] = [];
  for (const span of visibleSpans(text)) {
    const start = Math.max(index, span.index),
      end = Math.min(index + length, span.index + span.length);
    if (start >= end) continue;
    const next: [number, number, number] = [
        span.client,
        span.clock + start - span.index,
        end - start,
      ],
      tail = output.at(-1);
    if (tail && tail[0] === next[0] && tail[1] + tail[2] === next[1]) tail[2] += next[2];
    else output.push(next);
  }
  if (output.reduce((sum, range) => sum + range[2], 0) !== length)
    throw new OwnedTextError('Observed character range is outside its shared text.');
  return output;
}
export const inRanges = (ranges: readonly IdRange[], client: number, clock: number) =>
  ranges.some(
    ([owner, start, length]) => owner === client && start <= clock && clock < start + length
  );
export function ownershipCheck(doc: Y.Doc) {
  let links: Map<
    number,
    { clock: number; length: number; source: readonly [number, number] }[]
  > | null = null;
  return (ranges: readonly IdRange[], initialClient: number, initialClock: number): boolean => {
    let client = initialClient,
      clock = initialClock;
    const visited = new Set<string>();
    for (;;) {
      if (inRanges(ranges, client, clock)) return true;
      const key = `${client}:${clock}`;
      if (visited.has(key))
        throw new OwnedTextError('Restored character lineage contains a cycle.');
      visited.add(key);
      if (links === null) {
        links = new Map();
        const value = doc.share.get(OWNED_TEXT_LINEAGE);
        if (value instanceof Y.Map)
          for (const [name, data] of value) {
            if (!/^\d+:\d+$/.test(name))
              throw new OwnedTextError('Invalid restored character address.');
            const [owner, start] = name.split(':').map(Number),
              item = decodeLineage(data);
            if (!Number.isSafeInteger(owner) || !Number.isSafeInteger(start))
              throw new OwnedTextError('Invalid restored character clock.');
            const entries = links.get(owner!) ?? [];
            entries.push({ clock: start!, length: item.length, source: item.source });
            links.set(owner!, entries);
          }
        for (const entries of links.values()) entries.sort((a, b) => a.clock - b.clock);
      }
      const entries = links.get(client);
      if (!entries) return false;
      let left = 0,
        right = entries.length;
      while (left < right) {
        const middle = (left + right) >>> 1;
        if (entries[middle]!.clock <= clock) left = middle + 1;
        else right = middle;
      }
      const lineage = entries[left - 1];
      if (!lineage || clock >= lineage.clock + lineage.length) return false;
      client = lineage.source[0];
      clock = lineage.source[1] + clock - lineage.clock;
    }
  };
}
/** Undo/redo creates new items. Publish their ancestry before a native observer reads them. */
export function recordRestoredLineage(
  doc: Y.Doc,
  transaction: Y.Transaction,
  origin: symbol
): void {
  if (!(transaction.origin instanceof Y.UndoManager)) return;
  const entries: { key: string; value: string }[] = [];
  for (const structs of doc.store.clients.values())
    for (const item of structs) {
      if (!(item instanceof Y.Item) || !item.redone || !item.countable) continue;
      const target = item.redone,
        first = transaction.beforeState.get(target.client) ?? 0,
        last = transaction.afterState.get(target.client) ?? 0;
      if (target.clock < first || target.clock >= last) continue;
      entries.push({
        key: `${target.client}:${target.clock}`,
        value: JSON.stringify({ source: [item.id.client, item.id.clock], length: item.length }),
      });
    }
  if (entries.length)
    doc.transact(() => {
      const map = doc.getMap<string>(OWNED_TEXT_LINEAGE);
      for (const entry of entries) if (!map.has(entry.key)) map.set(entry.key, entry.value);
    }, origin);
}
