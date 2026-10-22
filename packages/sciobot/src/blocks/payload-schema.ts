import { jsonDataProblem, type JsonObject } from '../data/json';
export type { JsonObject, JsonValue } from '../data/json';

/** Native content controls carry the ID; a custom part carries this record. */
export interface BlockPayload extends JsonObject {
  readonly id: string;
  readonly type: string;
  readonly version: 1;
  readonly attributes: JsonObject;
}

export interface PayloadLimits {
  readonly maxBytes: number;
  readonly maxDepth: number;
  readonly maxNodes: number;
}

export const DEFAULT_PAYLOAD_LIMITS: PayloadLimits = Object.freeze({
  maxBytes: 2 * 1024 * 1024,
  maxDepth: 64,
  maxNodes: 50000,
});

export type PayloadDecodeResult =
  | { readonly status: 'supported'; readonly value: BlockPayload }
  | {
      readonly status: 'unsupported-version';
      readonly version: number;
      readonly raw: string;
      readonly value: JsonObject;
    }
  | { readonly status: 'invalid'; readonly reason: string; readonly raw: string };

export const BLOCK_TAG_PREFIX = 'scio:1:';
/** Review records share the application store but do not represent bound block payloads. */
export const PROPERTY_REVIEW_NODE_PREFIX = 'scio-review_';
export const MAX_BLOCK_ID_LENGTH = 64 - BLOCK_TAG_PREFIX.length;
const BLOCK_ID = /^[A-Za-z0-9][A-Za-z0-9_-]*$/;
const BLOCK_TYPE = /^[A-Za-z][A-Za-z0-9_.:-]{0,127}$/;
const own = (value: object, key: string) => Object.hasOwn(value, key);
const record = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value);

export function isBlockId(value: unknown): value is string {
  return typeof value === 'string' && value.length <= MAX_BLOCK_ID_LENGTH && BLOCK_ID.test(value);
}

function limitsOf(options: Partial<PayloadLimits>): PayloadLimits {
  const limits = { ...DEFAULT_PAYLOAD_LIMITS, ...options };
  for (const key of ['maxBytes', 'maxDepth', 'maxNodes'] as const) {
    if (
      !Number.isSafeInteger(limits[key]) ||
      limits[key] <= 0 ||
      limits[key] > DEFAULT_PAYLOAD_LIMITS[key]
    ) {
      throw new TypeError(`Invalid payload limit: ${key}`);
    }
  }
  return limits;
}

function schemaProblem(value: unknown): string | null {
  if (!record(value)) return 'expected-object';
  if (!own(value, 'id') || !isBlockId(value.id)) return 'invalid-id';
  if (!own(value, 'type') || typeof value.type !== 'string' || !BLOCK_TYPE.test(value.type))
    return 'invalid-type';
  if (!own(value, 'version') || !Number.isSafeInteger(value.version) || Number(value.version) < 1)
    return 'invalid-version';
  if (!own(value, 'attributes') || !record(value.attributes)) return 'invalid-attributes';
  return null;
}

function tooLarge(raw: string, limits: PayloadLimits): boolean {
  return raw.length > limits.maxBytes || new TextEncoder().encode(raw).byteLength > limits.maxBytes;
}

/** Unknown versions stay opaque. Callers retain their raw bytes and native fallback. */
export function decodeBlockPayload(
  raw: string,
  options: Partial<PayloadLimits> = {}
): PayloadDecodeResult {
  const limits = limitsOf(options);
  const invalid = (reason: string): PayloadDecodeResult => ({ status: 'invalid', reason, raw });
  if (tooLarge(raw, limits)) return invalid('byte-limit');
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return invalid('invalid-json');
  }
  const problem = jsonDataProblem(value, limits);
  if (problem) return invalid(problem);
  if (
    !record(value) ||
    !own(value, 'version') ||
    !Number.isSafeInteger(value.version) ||
    Number(value.version) < 1
  )
    return invalid('invalid-version');
  const object = value as JsonObject;
  const version = object.version as number;
  if (version !== 1) return { status: 'unsupported-version', version, raw, value: object };
  const schema = schemaProblem(value);
  if (schema) return invalid(schema);
  return { status: 'supported', value: object as BlockPayload };
}

export function encodeBlockPayload(
  value: BlockPayload,
  options: Partial<PayloadLimits> = {}
): string {
  const limits = limitsOf(options);
  const problem = jsonDataProblem(value, limits) ?? schemaProblem(value);
  if (problem || value.version !== 1) throw new TypeError(problem ?? 'unsupported-version');
  const raw = JSON.stringify(value);
  if (tooLarge(raw, limits)) throw new RangeError('byte-limit');
  return raw;
}
