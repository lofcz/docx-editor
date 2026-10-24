export const REVIEW_FLAG_TARGETS = 'sciobot-review-flag-targets-v1';
export const REVIEW_FLAG_INTENTS = 'sciobot-review-flag-intents-v1';
export const REVIEW_FLAG_RETRACTIONS = 'sciobot-review-flag-retractions-v1';
export const MAX_REVIEW_FLAGS = 50_000;
export const MAX_REVIEW_FLAG_INTENTS = 100_000;
export interface ReviewFlagTarget {
  readonly client: number;
  readonly clock: number;
  readonly key: string;
  readonly baseline: 'missing' | 'false';
}
export class ReviewFlagError extends Error {
  readonly code = 'invalid-review-flag-state';
}
const integer = (n: unknown): n is number => Number.isSafeInteger(n) && Number(n) >= 0;
const forbidden = new Set(['__proto__', 'prototype', 'constructor']);
export function validFlagKey(key: unknown): key is string {
  return (
    typeof key === 'string' &&
    key.length > 0 &&
    key.length <= 256 &&
    !forbidden.has(key) &&
    !/[\u0000-\u001f]/.test(key)
  );
}
export const flagAddress = (target: Pick<ReviewFlagTarget, 'client' | 'clock' | 'key'>) =>
  JSON.stringify([target.client, target.clock, target.key]);
export function decodeFlagTarget(address: string, value: unknown): ReviewFlagTarget {
  if (typeof value !== 'string' || value.length > 2048)
    throw new ReviewFlagError('Invalid bounded review flag target.');
  let data: unknown;
  try {
    data = JSON.parse(value);
  } catch {
    throw new ReviewFlagError('Invalid review flag target JSON.');
  }
  if (!data || typeof data !== 'object' || Array.isArray(data))
    throw new ReviewFlagError('Invalid review flag target.');
  const target = data as Record<string, unknown>;
  if (
    Object.keys(target).length !== 4 ||
    !Object.keys(target).every((key) => ['client', 'clock', 'key', 'baseline'].includes(key)) ||
    !integer(target.client) ||
    !integer(target.clock) ||
    !validFlagKey(target.key) ||
    !['missing', 'false'].includes(String(target.baseline))
  )
    throw new ReviewFlagError('Invalid review flag target fields.');
  const decoded = target as unknown as ReviewFlagTarget;
  if (flagAddress(decoded) !== address)
    throw new ReviewFlagError('Review flag address does not match its target.');
  return decoded;
}
export function validateFlagIntent(key: string, value: unknown): asserts value is string {
  if (
    key.length > 40 ||
    !/^\d+:\d+$/.test(key) ||
    !key.split(':').map(Number).every(integer) ||
    typeof value !== 'string' ||
    value.length > 2048
  )
    throw new ReviewFlagError('Invalid review flag intent.');
}

/** Exact immutable ownership evidence, supplied only by trusted native capture. */
export interface ReviewFlagReceipt {
  readonly id: string;
  readonly record: string;
}
export function decodeFlagRetraction(id: string, raw: unknown): ReviewFlagReceipt {
  if (typeof raw !== 'string' || raw.length > 8192)
    throw new ReviewFlagError('Invalid bounded flag retraction.');
  const item = JSON.parse(raw);
  if (
    !item ||
    typeof item !== 'object' ||
    Array.isArray(item) ||
    Object.keys(item).length !== 2 ||
    typeof item.id !== 'string' ||
    typeof item.record !== 'string'
  )
    throw new ReviewFlagError('Invalid flag retraction receipt.');
  validateFlagIntent(id, item.record);
  validateFlagIntent(item.id, item.record);
  return item;
}
