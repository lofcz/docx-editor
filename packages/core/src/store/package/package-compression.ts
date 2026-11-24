import { deflateInSlices, ZIP_WRITE_SLICE_BYTES, type ZipEntryCompressor } from './zip-write.ts';

// A document can retain compressed media across saves. Bound the extra storage per package.
const MAX_CACHE_BYTES = 16 * 1024 * 1024;
const compressors = new WeakMap<ReadonlyMap<string, Uint8Array>, ZipEntryCompressor>();

export function packageCompressor(parts: ReadonlyMap<string, Uint8Array>): ZipEntryCompressor {
  const existing = compressors.get(parts);
  if (existing) return existing;
  const entries = new WeakMap<Uint8Array, { source: Uint8Array; compressed: Uint8Array }>();
  const originals = new WeakSet(parts.values());
  let used = 0;
  function* compress(data: Uint8Array): Generator<void, Uint8Array> {
    if (data.length < ZIP_WRITE_SLICE_BYTES || !originals.has(data)) {
      return yield* deflateInSlices(data);
    }
    const hit = entries.get(data);
    // Part bytes can be visible to callers. Identity alone cannot prove they still match.
    if (hit) {
      let index = 0;
      while (index < data.length && data[index] === hit.source[index]) index++;
      if (index === data.length) return hit.compressed;
      used -= hit.source.length + hit.compressed.length;
      entries.delete(data);
    }
    if (data.length * 2 > MAX_CACHE_BYTES - used) return yield* deflateInSlices(data);
    const source = new Uint8Array(data);
    const compressed = yield* deflateInSlices(source);
    if (source.length + compressed.length <= MAX_CACHE_BYTES - used) {
      const prior = entries.get(data);
      if (prior) used -= prior.source.length + prior.compressed.length;
      entries.set(data, { source, compressed });
      used += source.length + compressed.length;
    }
    return compressed;
  }
  compressors.set(parts, compress);
  return compress;
}
