const typedArrayPrototype = Object.getPrototypeOf(Uint8Array.prototype) as object;
const readLength = Object.getOwnPropertyDescriptor(typedArrayPrototype, 'length')!.get!;
const readByteLength = Object.getOwnPropertyDescriptor(typedArrayPrototype, 'byteLength')!.get!;

/**
 * Copy genuine one-byte typed views into the Uint8Array resource contract.
 * Signed and clamped byte views preserve their eight-bit values at runtime.
 * Intrinsic access excludes caller-owned getters, iterators, and wider element types.
 */
export function copyImageBytes(value: unknown, maximum: number): Uint8Array {
  if (!ArrayBuffer.isView(value)) throw Error('Image resources require a genuine byte array.');
  const length = readLength.call(value) as unknown;
  const bytes = readByteLength.call(value) as unknown;
  if (
    typeof length !== 'number' ||
    length !== bytes ||
    !Number.isSafeInteger(length) ||
    length <= 0 ||
    length > maximum
  )
    throw Error('The image resource exceeds its byte bounds.');
  const copy = new Uint8Array(length);
  Uint8Array.prototype.set.call(copy, value as Uint8Array);
  return copy;
}
