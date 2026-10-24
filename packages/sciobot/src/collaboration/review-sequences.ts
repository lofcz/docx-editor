import { retainedBirthSequenceEntry } from './review-sequence-retention';
// Generic string sequence removals over pinned MIT Yjs item identities.
import * as Y from 'yjs';
import { ownershipCheck } from './ownership-positions';
import { OWNED_TEXT_LINEAGE } from './ownership-codec';
import { ReviewFieldError, type ReviewFieldTarget } from './review-field-codec';

export function reviewArray(doc: Y.Doc, target: ReviewFieldTarget): Y.Array<unknown> | null {
  if (target.root !== undefined)
    return doc.share.has(target.root) ? doc.getArray(target.root) : null;
  let id = Y.createID(target.client!, target.clock!);
  const seen = new Set<string>();
  for (let depth = 0; depth < 256; depth++) {
    const key = id.client + ':' + id.clock;
    if (seen.has(key)) throw new ReviewFieldError('Review sequence lineage contains a cycle.');
    seen.add(key);
    if (Y.getState(doc.store, id.client) <= id.clock) return null;
    const item = Y.getItem(doc.store, id);
    if (!(item instanceof Y.Item)) return null;
    if (item.redone) {
      id = item.redone;
      continue;
    }
    if (item.deleted) return null;
    if (!(item.content instanceof Y.ContentType) || !(item.content.type instanceof Y.Array))
      throw new ReviewFieldError('Review sequence target must name a shared array.');
    return item.content.type;
  }
  throw new ReviewFieldError('Review sequence lineage exceeds its limit.');
}

export function removedStringItems(array: Y.Array<unknown>, start: number, length: number) {
  const all: { key: string; value: unknown }[] = [];
  let at = 0;
  for (let item = array._start; item; item = item.right) {
    if (item.deleted || !item.countable) continue;
    const values = item.content.getContent();
    for (let offset = 0; offset < item.length; offset++, at++)
      all.push({
        key: item.id.client + ':' + (item.id.clock + offset),
        value: values[offset],
      });
  }
  return all.slice(start, start + length).flatMap((entry, offset) =>
    typeof entry.value === 'string'
      ? [
          {
            key: entry.key,
            value: entry.value,
            position: {
              left: all[start + offset - 1]?.key ?? null,
              right: all[start + offset + 1]?.key ?? null,
            },
          },
        ]
      : []
  );
}

export function originalReviewSequenceTarget(
  doc: Y.Doc,
  array: Y.Array<unknown>,
  key: string,
  targets: Iterable<ReviewFieldTarget>
) {
  const owns = ownershipCheck(doc),
    [client, clock] = key.split(':').map(Number);
  for (const target of targets) {
    if (!target.sequence || reviewArray(doc, target) !== array) continue;
    const [owner, start] = target.key.split(':').map(Number);
    if (owns([[owner!, start!, 1]], client!, clock!)) return target;
  }
  return null;
}

/** Normalize individual owned entries, preserving their original Yjs list position. */
export function normalizeReviewArray(
  doc: Y.Doc,
  array: Y.Array<unknown>,
  entries: readonly { target: ReviewFieldTarget; present: boolean; restoreOrder: boolean }[],
  remove: typeof Y.Array.prototype.delete,
  insert: typeof Y.Array.prototype.insert
) {
  if (array._item?.deleted) return;
  for (const { target, present, restoreOrder } of entries) {
    const owns = ownershipCheck(doc),
      [client, clock] = target.key.split(':').map(Number),
      ranges = [[client!, clock!, 1]] as const,
      visible: { index: number; client: number; clock: number }[] = [];
    const anchor = (key: string | null) => (key === null ? null : key.split(':').map(Number));
    const left = anchor(target.position!.left),
      right = anchor(target.position!.right);
    let leftAt: number | null = null,
      rightAt: number | null = null;
    let at = 0,
      position: number | null = null;
    for (let item = array._start; item; item = item.right) {
      const values = item.content.getContent();
      for (let offset = 0; offset < item.length; offset++) {
        if (owns(ranges, item.id.client, item.id.clock + offset)) {
          position ??= at;
          if (!item.deleted && item.countable) {
            if (values[offset] !== target.baseline)
              throw new ReviewFieldError('Restored sequence entry changed its value.');
            visible.push({ index: at, client: item.id.client, clock: item.id.clock + offset });
          }
        }
        if (!item.deleted && item.countable) at++;
        if (!item.deleted && item.countable) {
          if (left && owns([[left[0]!, left[1]!, 1]], item.id.client, item.id.clock + offset))
            leftAt = at;
          if (
            right &&
            rightAt === null &&
            owns([[right[0]!, right[1]!, 1]], item.id.client, item.id.clock + offset)
          )
            rightAt = at - 1;
        }
      }
    }
    if (at !== array.length)
      throw new ReviewFieldError('Review sequence visibility does not match its shared length.');
    const base = rightAt ?? (right === null ? at : (leftAt ?? (left === null ? 0 : position))),
      desired = base === null ? null : base - visible.filter((entry) => entry.index < base).length,
      keep =
        !present && !retainedBirthSequenceEntry(doc, array, target.baseline)
          ? null
          : visible
              .filter(
                (entry, index) =>
                  !target.initiallyPresent || !restoreOrder || entry.index - index === desired
              )
              .sort((a, b) => a.client - b.client || a.clock - b.clock)[0];
    // Undo restores new items. Their later normalization is outside the original redo
    // frame, so its final owner must also restore this single original list entry.
    if (array._searchMarker) array._searchMarker.length = 0;
    for (let i = visible.length - 1; i >= 0; i--)
      if (visible[i] !== keep) remove.call(array, visible[i]!.index, 1);
    if (present && !keep) {
      if (desired === null)
        throw new ReviewFieldError('Review sequence has no retained insertion position.');
      const restored = Y.getState(doc.store, doc.clientID);
      insert.call(array, desired, [target.baseline]);
      doc
        .getMap<string>(OWNED_TEXT_LINEAGE)
        .set(doc.clientID + ':' + restored, JSON.stringify({ source: [client, clock], length: 1 }));
    }
  }
}
