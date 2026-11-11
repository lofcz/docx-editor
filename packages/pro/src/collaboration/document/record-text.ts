/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import { NODE_ATOMIC_TEXT_FIELD, NODE_TEXT_FIELD, makeTextRecord } from './schema.ts';
import type { DocumentLimits } from './limits.ts';
import { ownedArraySpans } from './yjs-items.ts';

type TextLimits = Pick<DocumentLimits, 'maxTextLength' | 'maxChildren'>;

/** Each replacement owns one entry. Removing that entry reveals the previous live value. */
export function atomicTextState(
  record: Y.Map<unknown>,
  limits: TextLimits
): { value: string; entries: number; ownedEntries: number } | null | undefined {
  const history = record.get(NODE_ATOMIC_TEXT_FIELD);
  if (history === undefined) return undefined;
  if (!(history instanceof Y.Array) || history.length > limits.maxChildren) return null;
  let value = '';
  for (const entry of history) {
    if (typeof entry !== 'string' || entry.length > limits.maxTextLength) return null;
    value = entry;
  }
  const ownedEntries = history.doc
    ? ownedArraySpans(history, history.doc.clientID).reduce((count, span) => count + span.length, 0)
    : 0;
  return { value, entries: history.length, ownedEntries };
}

/** Generated fallback text retains whole replacement entries; ordinary text keeps characters. */
export function textRecord(value: string, atomic: boolean): Y.Map<unknown> {
  const record = makeTextRecord(atomic ? '' : value);
  if (atomic) {
    const history = new Y.Array<string>();
    history.insert(0, [value]);
    record.set(NODE_ATOMIC_TEXT_FIELD, history);
  }
  return record;
}

export function textRecordValue(record: Y.Map<unknown>, limits: TextLimits): string | null {
  const atomic = atomicTextState(record, limits);
  if (atomic !== undefined) return atomic?.value ?? null;
  const text = record.get(NODE_TEXT_FIELD);
  return text instanceof Y.Text ? text.toString() : '';
}

export function spliceAtomicText(
  record: Y.Map<unknown>,
  start: number,
  count: number,
  insert: string,
  atomic: boolean,
  limits: TextLimits
): boolean {
  if (!atomic && !record.has(NODE_ATOMIC_TEXT_FIELD)) return false;
  const before = textRecordValue(record, limits);
  if (before === null) throw Error('The atomic fallback text record is invalid.');
  const next = before.slice(0, start) + insert + before.slice(start + count);
  if (next.length > limits.maxTextLength)
    throw Error('The atomic fallback text exceeds the text limit.');
  let history = record.get(NODE_ATOMIC_TEXT_FIELD);
  if (history === undefined) {
    history = new Y.Array<string>();
    record.set(NODE_ATOMIC_TEXT_FIELD, history);
  }
  if (!(history instanceof Y.Array)) throw Error('The atomic replacement history is invalid.');
  const owned = history.doc ? ownedArraySpans(history, history.doc.clientID) : [];
  const removed = owned.reduce((count, span) => count + span.length, 0);
  if (history.length - removed >= limits.maxChildren)
    throw Error('The atomic replacement history exceeds the entry limit.');
  for (const span of owned.reverse()) history.delete(span.index, span.length);
  history.push([next]);
  return true;
}

/** Insert first for ordinary text; split slices retain their delete-first anchor behavior. */
export function spliceTextValue(
  text: Y.Text,
  start: number,
  count: number,
  insert: string,
  insertFirst: boolean
): void {
  if (insertFirst) {
    if (insert.length > 0) text.insert(start, insert);
    if (count > 0) text.delete(start + insert.length, count);
  } else {
    if (count > 0) text.delete(start, count);
    if (insert.length > 0) text.insert(start, insert);
  }
}
