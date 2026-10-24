// Provider-neutral structural attribute policy. The session supplies its own codec keys.
import type * as Y from 'yjs';
import { OWNED_TEXT_OPERATIONS, OWNED_TEXT_PREFIX, OwnedTextError } from './ownership-codec';
export const OWNED_STRUCTURE_FIELD = 'structural-attributes';
const empty: ReadonlySet<string> = new Set();
const decoded = new Map<string, ReadonlySet<string>>();
/** Conflicting headers remain a persisted refusal even when Yjs collects the losing header. */
export function assertOwnedProtocol(
  operations: Pick<Y.Map<unknown>, 'get' | '_map'> & { set(key: string, value: string): unknown },
  expected: string,
  allowMissing = false
): void {
  const value = operations.get('protocol');
  if (value === undefined && allowMissing) return;
  let item = operations._map.get('protocol'),
    count = 0,
    same = false,
    different = false;
  while (item) {
    if (++count > 4096) throw new OwnedTextError('Owned-text protocol history exceeds its limit.');
    for (const entry of item.content.getContent()) {
      if (entry === expected) same = true;
      else different = true;
    }
    item = item.left ?? undefined;
  }
  if (same && different) {
    operations.set('protocol', 'mixed');
    throw new OwnedTextError('Unsupported owned-text protocol.');
  }
  if (value !== expected) throw new OwnedTextError('Unsupported owned-text protocol.');
}

export function encodeStructuralPolicy(keys: readonly string[]): string {
  if (
    !Array.isArray(keys) ||
    keys.length > 64 ||
    keys.some(
      (key) =>
        typeof key !== 'string' ||
        !key.length ||
        key.length > 256 ||
        /[\u0000-\u001f]/.test(key) ||
        key.startsWith(OWNED_TEXT_PREFIX) ||
        ['__proto__', 'constructor', 'prototype'].includes(key)
    ) ||
    new Set(keys).size !== keys.length
  )
    throw new OwnedTextError('Invalid structural text attribute policy.');
  return JSON.stringify([...keys].sort());
}
export function structuralAttributes(
  doc: Y.Doc,
  expected: string | null,
  allowUnbound = false
): ReadonlySet<string> {
  const operations = doc.getMap(OWNED_TEXT_OPERATIONS),
    protocol = operations.get('protocol');
  const raw = operations.get(OWNED_STRUCTURE_FIELD);
  if (protocol !== '2') {
    if (raw !== undefined || expected !== null)
      throw new OwnedTextError('Structural attributes require owned-text protocol 2.');
    return empty;
  }
  if (raw === undefined) {
    if (!allowUnbound || expected !== null)
      throw new OwnedTextError('Bind a structural text attribute policy before editing.');
    return empty;
  }
  if (typeof raw !== 'string' || raw.length > 32768)
    throw new OwnedTextError('Invalid structural text attribute policy.');
  if (raw === 'mixed') throw new OwnedTextError('Incompatible structural text attribute policy.');
  let item = operations._map.get(OWNED_STRUCTURE_FIELD),
    count = 0;
  while (item) {
    if (++count > 4096)
      throw new OwnedTextError('Structural text policy history exceeds its limit.');
    if (item.content.getContent().some((value) => value !== raw)) {
      operations.set(OWNED_STRUCTURE_FIELD, 'mixed');
      throw new OwnedTextError('Incompatible structural text attribute policy.');
    }
    item = item.left ?? undefined;
  }
  if (expected !== null && expected !== raw)
    throw new OwnedTextError('Incompatible structural text attribute policy.');
  const cached = decoded.get(raw);
  if (cached) return cached;
  let keys: unknown;
  try {
    keys = JSON.parse(raw);
  } catch {
    throw new OwnedTextError('Invalid structural text attribute policy.');
  }
  if (
    !Array.isArray(keys) ||
    encodeStructuralPolicy(keys) !== raw ||
    (expected !== null && expected !== raw)
  )
    throw new OwnedTextError('Incompatible structural text attribute policy.');
  const result = new Set<string>(keys);
  if (decoded.size >= 64) decoded.clear();
  decoded.set(raw, result);
  return result;
}
/** Bind before editing. Mixed providers cannot silently assign different ownership semantics. */
export function bindStructuralAttributes(
  doc: Y.Doc,
  keys: readonly string[],
  previous: string | null
): string {
  const profile = encodeStructuralPolicy(keys),
    operations = doc.getMap<string>(OWNED_TEXT_OPERATIONS);
  if (operations.get('protocol') !== '2')
    throw new OwnedTextError('Structural attributes require owned-text protocol 2.');
  structuralAttributes(doc, previous, previous === null);
  const existing = operations.get(OWNED_STRUCTURE_FIELD);
  if (
    (previous !== null && profile !== previous) ||
    (existing !== undefined && existing !== profile)
  )
    throw new OwnedTextError('Incompatible structural text attribute policy.');
  if (existing === undefined) doc.transact(() => operations.set(OWNED_STRUCTURE_FIELD, profile));
  return profile;
}
