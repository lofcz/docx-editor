/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * The hidden places of characters a routed edit deletes.
 *
 * A paragraph can hold a character it does not show: a copy another paragraph shows, or the
 * original behind such a copy. Deleting only the shown place would let a hidden one show.
 */
import type * as Y from 'yjs';
import { deletePositions, itemIdsAt, type ItemId } from './yjs-items.ts';
import type { LogicalId } from './identity.ts';
import type { DocumentRegistry } from './registry.ts';
import type { Token } from './paragraph-text-diff.ts';
import type { ParagraphWrite } from './paragraph-text-moves.ts';
import { isCopyAt, textIdentities, type TextIdentities } from './paragraph-text-identity.ts';

/** What a write deletes from where its paragraph showed it, by identity. */
export function deletedIdentities(write: ParagraphWrite<Token>): Set<string> {
  const gone = new Set<string>();
  const retained = new Set(write.after.flatMap((token) => (token.source ? [token.source] : [])));
  for (const step of write.steps)
    if (step.op === 'eq') {
      const at = write.before[step.before]!.position;
      const identity = at === undefined ? null : write.identities[at];
      if (identity) retained.add(identity);
    }
  for (const step of write.steps) {
    if (step.op !== 'del') continue;
    const position = write.before[step.before]!.position;
    const identity = position === undefined ? null : write.identities[position];
    if (identity && !retained.has(identity)) gone.add(identity);
  }
  return gone;
}

/**
 * The hidden copies, in the plan's own paragraphs, of what the plan deletes. A paragraph can
 * hold a character twice, as when two peers join the same paragraphs at once, and show it
 * once: deleting the shown one would let the hidden one show in its place. Taken by item ID
 * before the writes, which shift positions; text a write moves gets new items and stays.
 */
export function hiddenCopiesInPlan<
  Write extends ParagraphWrite<Token> & {
    readonly text: Y.Text | null;
    readonly hidden: ReadonlySet<number>;
    readonly shownIdentities: TextIdentities | null;
  },
>(writes: readonly Write[]): HiddenCopies<Write>[] {
  const gone = new Set<string>();
  for (const write of writes) for (const identity of deletedIdentities(write)) gone.add(identity);
  if (gone.size === 0) return [];
  const copies: HiddenCopies<Write>[] = [];
  for (const write of writes) {
    if (!write.text || write.hidden.size === 0) continue;
    const positions = [...write.hidden].filter((position) => {
      const identity = write.identities[position];
      return identity !== null && identity !== undefined && gone.has(identity);
    });
    if (positions.length === 0) continue;
    const shown = write.shownIdentities;
    copies.push({
      write,
      ids: itemIdsAt(write.text, positions),
      originals: positions.map((position) =>
        shown && !shown.moved[position] && !shown.restored[position]
          ? (write.identities[position] ?? null)
          : null
      ),
    });
  }
  return copies;
}

/**
 * Hidden places of deleted characters in one planned paragraph, by item ID, and for each the
 * identity it holds in its first place, or null for a copy.
 */
export interface HiddenCopies<Write> {
  readonly write: Write;
  readonly ids: readonly ItemId[];
  readonly originals: readonly (string | null)[];
}

/**
 * Whether a held place of a deleted character goes. Every copy does: of two copies, the later
 * in document order shows. A character in its first place stays when the same edit copies it
 * again, because it hides behind that copy wherever the copy is, and a replica that receives
 * the deletion before the update the copy depends on would otherwise show neither.
 */
export function goes(original: string | null, copied: ReadonlySet<string>): boolean {
  return original === null || !copied.has(original);
}

/**
 * Delete, in paragraphs outside the plan, the characters a planned paragraph showed and no
 * longer shows: hidden copies would otherwise show in their place.
 */
export function deleteHiddenCopies(
  registry: DocumentRegistry,
  writes: readonly (ParagraphWrite<Token> & { readonly id: LogicalId })[],
  copied: ReadonlySet<string>
): void {
  const gone = new Set<string>();
  const planned = new Set<LogicalId>(writes.map((write) => write.id));
  for (const write of writes) for (const identity of deletedIdentities(write)) gone.add(identity);
  if (gone.size === 0) return;
  const holders = new Set<LogicalId>();
  for (const identity of gone) {
    for (const holder of registry.inline.follow.holdersOf(identity)) {
      if (!planned.has(holder)) holders.add(holder);
    }
  }
  for (const holder of holders) {
    const text = registry.inline.textOf(holder);
    if (!text) continue;
    // A character shows once, so every other place it is held is hidden.
    const identities = textIdentities(text);
    const doomed = identities.ids.flatMap((identity, position) => {
      if (!identity || !gone.has(identity)) return [];
      const copy = isCopyAt(identities, position);
      return goes(copy ? null : identity, copied) ? [position] : [];
    });
    if (doomed.length === 0) continue;
    deletePositions(text, doomed);
    registry.inline.follow.invalidate(holder);
    registry.inline.noteDrift();
  }
}
