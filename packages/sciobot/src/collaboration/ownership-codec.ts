import { copyBoundedJson } from '../data/json';

export const OWNED_TEXT_PROTOCOL = 1 as const;
export const OWNED_TEXT_PREFIX = 'sciobot-owned:v1:';
export const OWNED_TEXT_OPERATIONS = 'sciobot-owned-text-operations-v1';
export const OWNED_TEXT_LINEAGE = 'sciobot-owned-text-lineage-v1';
export type Attributes = Readonly<Record<string, unknown>>;
export type IdRange = readonly [client: number, clock: number, length: number];
export interface AttributeIntent {
  readonly kind: 'insert' | 'format';
  readonly stamp: readonly [number, number, number];
  readonly ranges: readonly IdRange[];
  readonly values: Attributes;
}
export interface Lineage {
  readonly source: readonly [number, number];
  readonly length: number;
}
export class OwnedTextError extends Error {
  readonly code = 'invalid-owned-text-state';
}
const limits = { maxBytes: 16 * 1024 * 1024, maxNodes: 1_000_000, maxDepth: 32 };
const number = (value: unknown): value is number =>
  Number.isSafeInteger(value) && Number(value) >= 0;
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const known = (value: Record<string, unknown>, keys: readonly string[]) =>
  Object.keys(value).every((key) => keys.includes(key));
const forbidden = new Set(['__proto__', 'prototype', 'constructor']);
function safeKeys(value: unknown): boolean {
  if (value === null || typeof value !== 'object') return true;
  return Object.entries(value).every(([key, child]) => !forbidden.has(key) && safeKeys(child));
}
export function detachedAttributes(input: unknown): Record<string, unknown> {
  // Reserve space in each record for observed character ranges and its envelope.
  const copied = copyBoundedJson(input, { ...limits, maxBytes: 8 * 1024 * 1024 });
  if (!copied.ok || !record(copied.value) || !safeKeys(copied.value))
    throw new OwnedTextError('Owned attributes need bounded JSON data.');
  const output: Record<string, unknown> = Object.create(null);
  for (const [key, value] of Object.entries(copied.value)) {
    if (key.startsWith(OWNED_TEXT_PREFIX)) continue;
    output[key] = value;
  }
  return output;
}
export function encodeIntent(value: AttributeIntent): string {
  const result = JSON.stringify(value);
  if (new TextEncoder().encode(result).length > limits.maxBytes)
    throw new OwnedTextError('Owned attribute intent exceeds its byte limit.');
  return result;
}
export function decodeIntent(input: unknown): AttributeIntent {
  if (typeof input !== 'string' || input.length > limits.maxBytes)
    throw new OwnedTextError('Invalid owned attribute record.');
  let data: unknown;
  try {
    data = JSON.parse(input);
  } catch {
    throw new OwnedTextError('Invalid owned attribute JSON.');
  }
  const copied = copyBoundedJson(data, limits);
  if (!copied.ok || !record(copied.value))
    throw new OwnedTextError('Invalid bounded attribute record.');
  const item = copied.value;
  if (
    !known(item, ['kind', 'stamp', 'ranges', 'values']) ||
    !['insert', 'format'].includes(String(item.kind)) ||
    !Array.isArray(item.stamp) ||
    item.stamp.length !== 3 ||
    !item.stamp.every(number) ||
    !Array.isArray(item.ranges) ||
    !item.ranges.length ||
    item.ranges.length > 100_000 ||
    !item.ranges.every(
      (range) =>
        Array.isArray(range) &&
        range.length === 3 &&
        range.every(number) &&
        Number(range[2]) > 0 &&
        Number(range[1]) + Number(range[2]) <= Number.MAX_SAFE_INTEGER
    ) ||
    !record(item.values) ||
    !safeKeys(item.values) ||
    Object.keys(item.values).some((key) => key.startsWith(OWNED_TEXT_PREFIX))
  )
    throw new OwnedTextError('Invalid owned attribute schema.');
  return item as unknown as AttributeIntent;
}
export function decodeLineage(input: unknown): Lineage {
  if (typeof input !== 'string' || input.length > 1024)
    throw new OwnedTextError('Invalid restored character lineage.');
  let item: unknown;
  try {
    item = JSON.parse(input);
  } catch {
    throw new OwnedTextError('Invalid restored character JSON.');
  }
  if (
    !record(item) ||
    !known(item, ['source', 'length']) ||
    !Array.isArray(item.source) ||
    item.source.length !== 2 ||
    !item.source.every(number) ||
    !number(item.length) ||
    item.length === 0 ||
    Number(item.source[1]) + item.length > Number.MAX_SAFE_INTEGER
  )
    throw new OwnedTextError('Invalid restored character schema.');
  return item as unknown as Lineage;
}
export function nativeAttributes(input: Attributes): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(input).filter(([key]) => !key.startsWith(OWNED_TEXT_PREFIX))
  );
}
export const compareIntents = (left: AttributeIntent, right: AttributeIntent) =>
  left.stamp[0] - right.stamp[0] ||
  left.stamp[1] - right.stamp[1] ||
  left.stamp[2] - right.stamp[2];
