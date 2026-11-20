/**
 * A `ReadonlyMap` that is one at RUNTIME, not only to the type checker.
 *
 * Reads delegate to the source map; there is no `set`, `delete` or `clear` to find on the
 * object at all, so a consumer that casts the type away still cannot write through it.
 */
export function readonlyMapView<K, V>(source: Map<K, V>): ReadonlyMap<K, V> {
  const view = Object.freeze({
    get: (key: K) => source.get(key),
    has: (key: K) => source.has(key),
    forEach: (callback: (value: V, key: K, map: ReadonlyMap<K, V>) => void, thisArg?: unknown) => {
      for (const [key, value] of source) callback.call(thisArg, value, key, view);
    },
    keys: () => source.keys(),
    values: () => source.values(),
    entries: () => source.entries(),
    [Symbol.iterator]: () => source[Symbol.iterator](),
    [Symbol.toStringTag]: 'Map',
    get size() {
      return source.size;
    },
  }) as ReadonlyMap<K, V>;
  return view;
}
