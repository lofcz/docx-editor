/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { DocumentRegistry } from './registry.ts';
import { asLogicalId, type LogicalId } from './identity.ts';
import { INLINE_FIELD } from './paragraph-text.ts';
import { MAX_MARK_CHAIN, marksOfText, type MarkLookup } from './paragraph-text-marks.ts';
import { fieldOf } from './yjs-items.ts';

/** A restored copy keeps the relocation history of the character it restores. */
function relocated(marks: MarkLookup, client: number, clock: number): boolean {
  for (let depth = 0; depth < MAX_MARK_CHAIN; depth += 1) {
    if (marks.origin('move', client, clock)) return true;
    const restored = marks.origin('restore', client, clock);
    if (!restored) return false;
    const next = restored.clock + clock - restored.firstClock;
    if (restored.client === client && next === clock) return true;
    client = restored.client;
    clock = next;
  }
  return true;
}

/**
 * Independent paragraph births own fresh content. Enter moves existing content instead.
 * Inspect only the history step's items, including deleted ones: a peer's later relocation
 * into a fresh paragraph must not turn that paragraph into the creator's split operation.
 */
export function independentParagraphBirths(
  registry: DocumentRegistry,
  born: ReadonlySet<LogicalId>,
  changed: readonly Y.Item[]
): ReadonlySet<LogicalId> {
  const texts = new Map<Y.Text, boolean>();
  for (const item of changed) {
    const text = item.parent;
    if (!(text instanceof Y.Text) || fieldOf(text) !== INLINE_FIELD || !item.countable) continue;
    let moved = texts.get(text) ?? false;
    const marks = marksOfText(text);
    for (let clock = item.id.clock; !moved && clock < item.id.clock + item.length; clock += 1)
      moved = relocated(marks, item.id.client, clock);
    if (!moved && !(item.content instanceof Y.ContentString))
      for (const value of item.content.getContent())
        if (
          value &&
          typeof value === 'object' &&
          'n' in value &&
          typeof value.n === 'string' &&
          !born.has(asLogicalId(value.n))
        )
          moved = true;
    texts.set(text, moved);
  }
  const independent = new Set<LogicalId>();
  for (const id of born) {
    if (registry.kindOf(id) !== 'paragraph') continue;
    const text = registry.inline.textOf(id);
    if (text && texts.get(text) === false) independent.add(id);
  }
  return independent;
}
