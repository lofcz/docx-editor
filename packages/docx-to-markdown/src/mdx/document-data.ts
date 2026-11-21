import type { JsonValue } from './registry';

export const MAX_SOURCE_BYTES = 2 * 1024 * 1024;
const MAX_NODES = 100_000;
const MAX_DEPTH = 64;

/** Drop parser coordinates and compiler trees. Preserve the original expression text as data. */
export function syntaxData(input: unknown): JsonValue {
  let visited = 0;
  function copy(value: unknown, depth: number): JsonValue {
    if (++visited > MAX_NODES || depth > MAX_DEPTH)
      throw new RangeError('MDX tree exceeds the document limit.');
    if (typeof value === 'string' || typeof value === 'boolean' || value === null) return value;
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (Array.isArray(value)) return Object.freeze(value.map((item) => copy(item, depth + 1)));
    if (typeof value !== 'object' || value === null)
      throw new TypeError('Invalid MDX syntax data.');
    const result: Record<string, JsonValue> = Object.create(null);
    for (const [key, item] of Object.entries(value)) {
      if (key === 'data' || key === 'position' || item === undefined) continue;
      if (key === '__proto__' || key === 'constructor' || key === 'prototype')
        throw new TypeError('Invalid MDX syntax key.');
      result[key] = copy(item, depth + 1);
    }
    return Object.freeze(result);
  }
  return copy(input, 0);
}
