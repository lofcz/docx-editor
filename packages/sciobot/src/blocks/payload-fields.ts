import type { OoxmlElement, OoxmlNode } from '@docx-editor.dev/core/store';
import type { JsonObject, JsonValue } from './payload-schema';
import { DEFAULT_PAYLOAD_LIMITS } from './payload-schema';

/**
 * Payload field registers.
 *
 * A record's `data` text holds the payload as it was created. Shared text merges character by
 * character, so two sessions rewriting that JSON at once produce a third string that neither
 * wrote: a garbled value, or JSON that no longer parses and makes the whole store unreadable.
 *
 * Later writes therefore leave `data` alone and put each changed field on a fresh child
 * element. Prior elements stay in document order. Concurrent replacements keep separate
 * ownership, so undo cannot remove another session's value or change its ordering anchor.
 * Disjoint fields live on different elements, so they both survive.
 *
 * `<attr name="k" value="json"/>` overrides `attributes[k]`; `<field name="k" value="json"/>`
 * overrides a top-level key. An element without `value` removes the key. When concurrent first
 * writes leave two elements for one key, the last in document order wins in every session.
 * The existing field-count limit bounds retained replacements; exceeding it refuses the write.
 */
export const ATTRIBUTE_FIELD = 'attr';
export const TOP_FIELD = 'field';
const NAME = 'name';
const VALUE = 'value';
/** Identity and schema version never change after creation; `attributes` is split per key. */
const FIXED_KEYS = new Set(['id', 'version', 'attributes']);
const MAX_FIELDS = 10_000;
const MAX_FIELD_BYTES = 8 * 1024 * 1024;

type FieldKind = typeof ATTRIBUTE_FIELD | typeof TOP_FIELD;
type Overrides = Map<string, JsonValue | undefined>;

function plainAttribute(element: OoxmlElement, localName: string): string | undefined {
  return element.attributes.find(
    (attribute) => attribute.localName === localName && attribute.namespaceUri === ''
  )?.value;
}

function fieldKind(node: OoxmlNode, namespaceUri: string): FieldKind | null {
  if (node.kind === 'textValue' || node.namespaceUri !== namespaceUri) return null;
  return node.localName === ATTRIBUTE_FIELD || node.localName === TOP_FIELD ? node.localName : null;
}

export type FieldRead =
  | { readonly ok: true; readonly count: number; readonly apply: (base: unknown) => unknown }
  | { readonly ok: false; readonly reason: string };

/** Overrides in document order. Pure; the caller validates the merged payload. */
export function readFields(record: OoxmlElement): FieldRead {
  const top: Overrides = new Map();
  const attributes: Overrides = new Map();
  let count = 0;
  let bytes = 0;
  for (const child of record.children) {
    const kind = fieldKind(child, record.namespaceUri);
    if (!kind) continue;
    const element = child as OoxmlElement;
    if (++count > MAX_FIELDS) return { ok: false, reason: 'field-limit' };
    const name = plainAttribute(element, NAME);
    if (!name || (kind === TOP_FIELD && FIXED_KEYS.has(name)))
      return { ok: false, reason: 'invalid-field-name' };
    const raw = plainAttribute(element, VALUE);
    let value: JsonValue | undefined;
    if (raw !== undefined) {
      if (raw.length > DEFAULT_PAYLOAD_LIMITS.maxBytes)
        return { ok: false, reason: 'field-byte-limit' };
      const valueBytes = new TextEncoder().encode(raw).byteLength;
      if (valueBytes > DEFAULT_PAYLOAD_LIMITS.maxBytes)
        return { ok: false, reason: 'field-byte-limit' };
      bytes += valueBytes;
      if (bytes > MAX_FIELD_BYTES) return { ok: false, reason: 'field-byte-limit' };
      try {
        value = JSON.parse(raw) as JsonValue;
      } catch {
        return { ok: false, reason: 'invalid-field-value' };
      }
    }
    (kind === TOP_FIELD ? top : attributes).set(name, value);
  }
  const apply = (base: unknown): unknown => {
    if (!count || typeof base !== 'object' || base === null || Array.isArray(base)) return base;
    const merged: Record<string, unknown> = { ...(base as Record<string, unknown>) };
    for (const [name, value] of top) {
      if (value === undefined) delete merged[name];
      else
        Object.defineProperty(merged, name, {
          value,
          enumerable: true,
          configurable: true,
          writable: true,
        });
    }
    if (attributes.size) {
      const current = merged.attributes;
      if (typeof current !== 'object' || current === null || Array.isArray(current)) return merged;
      const next: Record<string, unknown> = { ...(current as Record<string, unknown>) };
      for (const [name, value] of attributes) {
        if (value === undefined) delete next[name];
        else
          Object.defineProperty(next, name, {
            value,
            enumerable: true,
            configurable: true,
            writable: true,
          });
      }
      merged.attributes = next;
    }
    return merged;
  };
  return { ok: true, count, apply };
}

/** JSON with sorted object keys, so equal values compare equal regardless of key order. */
function stable(value: JsonValue | undefined): string {
  if (value === undefined) return '';
  if (typeof value !== 'object' || value === null) return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map((item) => stable(item)).join(',')}]`;
  const object = value as JsonObject;
  const keys = Object.keys(object).sort();
  return `{${keys.map((key) => `${JSON.stringify(key)}:${stable(object[key])}`).join(',')}}`;
}

interface FieldChange {
  readonly kind: FieldKind;
  readonly name: string;
  readonly value: JsonValue | undefined;
}

function changes(before: JsonObject, after: JsonObject, kind: FieldKind): FieldChange[] {
  const found: FieldChange[] = [];
  const names = new Set([...Object.keys(before), ...Object.keys(after)]);
  for (const name of names) {
    if (kind === TOP_FIELD && FIXED_KEYS.has(name)) continue;
    const next = Object.hasOwn(after, name) ? after[name] : undefined;
    const previous = Object.hasOwn(before, name) ? before[name] : undefined;
    if (stable(previous) !== stable(next)) found.push({ kind, name, value: next });
  }
  return found;
}

/** Every changed field between two payloads with the same identity. */
export function fieldChanges(before: JsonObject, after: JsonObject): FieldChange[] {
  return [
    ...changes(before, after, TOP_FIELD),
    ...changes(before.attributes as JsonObject, after.attributes as JsonObject, ATTRIBUTE_FIELD),
  ];
}

function fieldElement(id: string, namespaceUri: string, change: FieldChange): OoxmlElement {
  const attribute = (localName: string, value: string) => ({
    kind: 'genericExtension' as const,
    namespaceUri: '',
    localName,
    value,
  });
  return {
    id,
    kind: 'generic',
    namespaceUri,
    localName: change.kind,
    namespaceBindings: [],
    children: [],
    attributes: [
      attribute(NAME, change.name),
      ...(change.value === undefined ? [] : [attribute(VALUE, JSON.stringify(change.value))]),
    ],
  } as OoxmlElement;
}

/** Compact validated records only before a room starts. Never remove opaque XML metadata. */
export function compactPreparedFields(record: OoxmlElement): OoxmlElement {
  const latest = new Map<string, OoxmlNode>();
  for (const child of record.children) {
    const kind = fieldKind(child, record.namespaceUri);
    if (kind) latest.set(`${kind}:${plainAttribute(child as OoxmlElement, NAME)}`, child);
  }
  const children = record.children.filter((child) => {
    const kind = fieldKind(child, record.namespaceUri);
    if (!kind || child.kind === 'textValue') return true;
    if (latest.get(`${kind}:${plainAttribute(child, NAME)}`) === child) return true;
    return (
      child.children.length > 0 ||
      child.namespaceBindings.length > 0 ||
      child.attributes.some(
        (attribute) =>
          attribute.namespaceUri !== '' ||
          (attribute.localName !== NAME && attribute.localName !== VALUE)
      )
    );
  });
  return children.length === record.children.length
    ? record
    : ({ ...record, children } as OoxmlElement);
}

/**
 * The record's children after these changes. A changed field gets a fresh element identity.
 * Unknown attributes and children remain intact. Other fields keep their existing identities.
 */
export function withFieldChanges(
  record: OoxmlElement,
  updates: readonly FieldChange[],
  usedIds: ReadonlySet<string>
): OoxmlNode[] {
  const children = [...record.children];
  const freshId = () => {
    let id: string;
    // Removed IDs still exist in shared history and must never identify another replacement.
    do id = `field_${crypto.randomUUID().replaceAll('-', '')}`;
    while (usedIds.has(id) || children.some((child) => child.id === id));
    return id;
  };
  for (const change of updates) {
    const replacement = fieldElement(freshId(), record.namespaceUri, change);
    children.push(replacement);
  }
  return children;
}
