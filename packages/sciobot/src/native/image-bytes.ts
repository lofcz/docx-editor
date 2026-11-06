const typedArrayPrototype = Object.getPrototypeOf(Uint8Array.prototype) as object;
const readLength = Object.getOwnPropertyDescriptor(typedArrayPrototype, 'length')!.get!;
const readByteLength = Object.getOwnPropertyDescriptor(typedArrayPrototype, 'byteLength')!.get!;

/** Authenticate and copy bounded bytes across host realms without caller-owned getters or iterators. */
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
