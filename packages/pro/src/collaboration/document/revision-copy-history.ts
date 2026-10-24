/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import { asLogicalId, parseClientClock } from './identity.ts';
import type { DocumentRegistry } from './registry.ts';
import { revisionCopyPositions } from './revision-text-copies.ts';
import { textIdentities } from './paragraph-text-identity.ts';
import { deleteSetRuns, fieldOf, holderItem, structAt } from './yjs-items.ts';
import { RESTORE_ORIGIN } from './paragraph-text-restore.ts';

export interface RevisionCopyHistory {
  readonly paragraph: string;
  readonly fields: readonly {
    readonly key: string;
    readonly value: string;
    readonly items: readonly string[];
  }[];
}

/** Capture only paired revision keys whose format markers this exact history step introduced. */
export function captureRevisionCopyHistory(
  registry: DocumentRegistry,
  step: unknown
): readonly RevisionCopyHistory[] {
  const keys = new Map<Y.Text, Set<string>>();
  const inserted = (step as { insertions?: unknown } | undefined)?.insertions;
  for (const run of deleteSetRuns(inserted)) {
    for (let clock = run.clock; clock < run.clock + run.length; ) {
      const item = structAt(registry.doc, { client: run.client, clock });
      if (!item) break;
      if (
        item.parent instanceof Y.Text &&
        item.content instanceof Y.ContentFormat &&
        item.content.key.startsWith('v:') &&
        item.content.value !== null
      ) {
        const set = keys.get(item.parent) ?? new Set<string>();
        set.add(item.content.key);
        keys.set(item.parent, set);
      }
      clock = item.id.clock + item.length;
    }
  }
  const result: RevisionCopyHistory[] = [];
  for (const [text, owned] of keys) {
    const record = holderItem(text)?.parent;
    const paragraph = record instanceof Y.Map ? fieldOf(record) : null;
    if (!paragraph || registry.inline.textOf(asLogicalId(paragraph)) !== text) continue;
    const ids = textIdentities(text),
      copies = revisionCopyPositions(text, ids, registry.limits, paragraph);
    const fields = new Map<string, { key: string; value: string; items: string[] }>();
    for (const [position, copy] of copies) {
      if (!owned.has(copy.role.key)) continue;
      const value = copy.attributes[copy.role.key];
      if (typeof value !== 'string') continue;
      const entry = fields.get(copy.role.key) ?? { key: copy.role.key, value, items: [] };
      entry.items.push(ids.items[position]!);
      fields.set(copy.role.key, entry);
    }
    if (fields.size) result.push({ paragraph, fields: [...fields.values()] });
  }
  return result;
}

/** Follow author-local redo metadata only while writing that author's explicit history restoration. */
function restoredItem(doc: Y.Doc, itemId: string): string | null {
  let id = parseClientClock(itemId);
  for (let depth = 0; id && depth < 64; depth++) {
    const item = structAt(doc, id);
    if (!item) return null;
    if (!item.deleted) return `${id.client}:${id.clock}`;
    if (!item.redone) return null;
    id = { client: item.redone.client, clock: item.redone.clock + id.clock - item.id.clock };
  }
  return null;
}

/** Undo withdraws only this step's paired keys; redo restores their observed character cohorts. */
export function restoreRevisionCopyHistory(
  registry: DocumentRegistry,
  snapshots: readonly RevisionCopyHistory[],
  direction: 'undo' | 'redo'
): void {
  if (!snapshots.length) return;
  registry.doc.transact(() => {
    for (const snapshot of snapshots) {
      const paragraph = asLogicalId(snapshot.paragraph),
        text = registry.inline.textOf(paragraph);
      if (!text) continue;
      const ids = textIdentities(text),
        positions = new Map(ids.items.map((item, index) => [item, index]));
      for (const field of snapshot.fields) {
        if (direction === 'undo') {
          text.format(0, text.length, { [field.key]: null });
          continue;
        }
        const selected = field.items
          .flatMap((item) => {
            const restored = restoredItem(registry.doc, item),
              at = restored ? positions.get(restored) : undefined;
            return at === undefined ? [] : [at];
          })
          .sort((a, b) => a - b);
        for (let at = selected.length - 1; at >= 0; ) {
          const end = selected[at]!;
          let start = end;
          at--;
          while (at >= 0 && selected[at] === start - 1) {
            start = selected[at]!;
            at--;
          }
          text.format(start, end - start + 1, { [field.key]: field.value });
        }
      }
      registry.inline.follow.invalidate(paragraph);
    }
  }, RESTORE_ORIGIN);
}
