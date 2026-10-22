export const REVIEW_FIELD_TARGETS = 'sciobot-review-field-targets-v1';
export const REVIEW_FIELD_INTENTS = 'sciobot-review-field-intents-v1';
export const MAX_REVIEW_FIELD_TARGETS = 50000;
export const MAX_REVIEW_FIELD_INTENTS = 100000;
export const MAX_REVIEW_FIELD_BYTES = 8 * 1024 * 1024;
export interface ReviewFieldLocation {
  /** String sequence entries use the item address in key and own only their removal. */
  readonly sequence?: true;
  readonly root?: string;
  readonly client?: number;
  readonly clock?: number;
  readonly key: string;
}
export interface ReviewFieldTarget extends ReviewFieldLocation {
  readonly baseline: string | null;
  readonly initiallyPresent?: boolean;
  readonly position?: { readonly left: string | null; readonly right: string | null };
}
export interface ReviewFieldIntent {
  readonly target: string;
  readonly value: string | null;
  readonly stamp: readonly [number, number, number];
}
export class ReviewFieldError extends Error {
  readonly code = 'invalid-review-field-state';
}
export const fieldAddress = (target: ReviewFieldLocation) =>
  JSON.stringify(
    target.root !== undefined
      ? [target.sequence ? 'array-root' : 'root', target.root, target.key]
      : [target.sequence ? 'array-item' : 'item', target.client, target.clock, target.key]
  );
export function encodeFieldRecord(value: unknown) {
  const encoded = JSON.stringify(value);
  if (new TextEncoder().encode(encoded).length > MAX_REVIEW_FIELD_BYTES)
    throw new ReviewFieldError('Owned review field data exceeds its byte limit.');
  return encoded;
}
function parse(raw: unknown): Record<string, unknown> {
  if (typeof raw !== 'string' || new TextEncoder().encode(raw).length > MAX_REVIEW_FIELD_BYTES)
    throw new ReviewFieldError('Invalid bounded review field data.');
  const value = JSON.parse(raw);
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new ReviewFieldError('Invalid review field record.');
  return value;
}
const integer = (v: unknown): v is number => Number.isSafeInteger(v) && Number(v) >= 0;
export function decodeFieldTarget(key: string, raw: unknown): ReviewFieldTarget {
  const value = parse(raw),
    keys = Object.keys(value);
  if (
    typeof value.key !== 'string' ||
    value.key.length > 65536 ||
    !value.key.length ||
    !(value.baseline === null || typeof value.baseline === 'string') ||
    (value.sequence !== undefined && value.sequence !== true) ||
    keys.some(
      (k) =>
        ![
          'sequence',
          'root',
          'client',
          'clock',
          'key',
          'baseline',
          'position',
          'initiallyPresent',
        ].includes(k)
    ) ||
    ((value.position !== undefined || value.initiallyPresent !== undefined) &&
      value.sequence !== true)
  )
    throw new ReviewFieldError('Invalid review field target.');
  if (value.root !== undefined) {
    if (
      typeof value.root !== 'string' ||
      value.root.length > 65536 ||
      !value.root.length ||
      keys.length !== 3 + 3 * Number(value.sequence === true)
    )
      throw new ReviewFieldError('Invalid review root field.');
  } else if (
    !integer(value.client) ||
    !integer(value.clock) ||
    keys.length !== 4 + 3 * Number(value.sequence === true)
  )
    throw new ReviewFieldError('Invalid review item field.');
  if (
    value.sequence &&
    (typeof value.baseline !== 'string' ||
      typeof value.initiallyPresent !== 'boolean' ||
      !/^\d+:\d+$/.test(value.key) ||
      !value.key.split(':').map(Number).every(integer))
  )
    throw new ReviewFieldError('Invalid review sequence address.');
  if (value.sequence) {
    const position = value.position as ReviewFieldTarget['position'];
    if (
      !position ||
      typeof position !== 'object' ||
      Object.keys(position).length !== 2 ||
      !['left', 'right'].every((key) => Object.hasOwn(position, key)) ||
      ![position.left, position.right].every(
        (key) =>
          key === null ||
          (typeof key === 'string' &&
            key.length <= 40 &&
            /^\d+:\d+$/.test(key) &&
            key.split(':').map(Number).every(integer))
      )
    )
      throw new ReviewFieldError('Invalid review sequence position.');
  }
  const target = value as unknown as ReviewFieldTarget;
  if (fieldAddress(target) !== key)
    throw new ReviewFieldError('Review field address does not match its target.');
  return target;
}
export function decodeFieldIntent(key: string, raw: unknown): ReviewFieldIntent {
  const value = parse(raw);
  if (
    !/^\d+:\d+$/.test(key) ||
    key.length > 40 ||
    !key.split(':').map(Number).every(integer) ||
    Object.keys(value).length !== 3 ||
    typeof value.target !== 'string' ||
    value.target.length > 262144 ||
    !(value.value === null || typeof value.value === 'string') ||
    !Array.isArray(value.stamp) ||
    value.stamp.length !== 3 ||
    !value.stamp.every(integer)
  )
    throw new ReviewFieldError('Invalid review field intent.');
  return value as unknown as ReviewFieldIntent;
}
export function newerFieldIntent(a: ReviewFieldIntent, b: ReviewFieldIntent) {
  for (let i = 0; i < 3; i++) if (a.stamp[i] !== b.stamp[i]) return a.stamp[i]! > b.stamp[i]!;
  return false;
}
