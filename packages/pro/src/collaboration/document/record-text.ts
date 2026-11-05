/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import { NODE_ATOMIC_TEXT_FIELD, NODE_TEXT_FIELD, makeTextRecord } from './schema.ts';

/** Generated fallback text replaces one scalar; ordinary text keeps character ownership. */
export function textRecord(value: string, atomic: boolean): Y.Map<unknown> {
  const record = makeTextRecord(atomic ? '' : value);
  if (atomic) record.set(NODE_ATOMIC_TEXT_FIELD, value);
  return record;
}

export function textRecordValue(record: Y.Map<unknown>): string | null {
  const atomic = record.get(NODE_ATOMIC_TEXT_FIELD);
  if (atomic !== undefined) return typeof atomic === 'string' ? atomic : null;
  const text = record.get(NODE_TEXT_FIELD);
  return text instanceof Y.Text ? text.toString() : '';
}

export function spliceAtomicText(
  record: Y.Map<unknown>,
  start: number,
  count: number,
  insert: string,
  atomic: boolean
): boolean {
  if (!atomic && !record.has(NODE_ATOMIC_TEXT_FIELD)) return false;
  const before = textRecordValue(record);
  if (before === null) throw Error('The atomic fallback text record is invalid.');
  record.set(NODE_ATOMIC_TEXT_FIELD, before.slice(0, start) + insert + before.slice(start + count));
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
