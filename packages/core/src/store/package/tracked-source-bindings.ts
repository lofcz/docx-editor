import { isXmlNCName } from './qname.ts';
import type { OoxmlElement, OoxmlNode } from './ooxml-tree.ts';

/** Load-local native identities. Serialization consumes these after transferring their proofs. */
export const TRACKED_SOURCE_BINDINGS_NS = 'urn:docx-editor:tracked-source-bindings:1';
const ORIGIN_NS = 'urn:docx-editor:tracked-run-origins:1';
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const MAX_BYTES = 2 * 1024 * 1024;
const CHUNK_LENGTH = 2048;
export interface TrackedSourceBinding {
  readonly scope: string;
  readonly runId: string;
  readonly token: string;
  readonly absent: boolean;
  readonly retained: readonly string[];
  readonly texts: readonly {
    readonly id: string;
    readonly token: string;
    readonly space: string;
  }[];
}
const object = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const keys = (value: object, names: readonly string[]) =>
  Object.keys(value).length === names.length &&
  Object.keys(value).every((key) => names.includes(key));
const id = (value: unknown): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= 4096;
const token = (value: unknown, scope: string): value is string =>
  typeof value === 'string' &&
  value.startsWith(scope + '/') &&
  UUID.test(value.slice(scope.length + 1));
function invalid(): never {
  throw new Error('Invalid tracked source bindings.');
}
export function readTrackedSourceBindings(node: OoxmlNode): readonly TrackedSourceBinding[] {
  if (node.kind !== 'paragraph') return [];
  const attrs = node.attributes.filter((attr) => attr.namespaceUri === TRACKED_SOURCE_BINDINGS_NS);
  if (!attrs.length) return [];
  let encoded: string;
  if (attrs.length === 1 && attrs[0]!.localName === 'bindings') encoded = attrs[0]!.value;
  else {
    if (
      attrs.length > 1024 ||
      attrs.some((attribute) => !/^bindings(?:0|[1-9][0-9]*)$/.test(attribute.localName))
    )
      invalid();
    const chunks = new Map(
      attrs.map((attribute) => [
        Number(attribute.localName.slice('bindings'.length)),
        attribute.value,
      ])
    );
    if (chunks.size !== attrs.length) invalid();
    const ordered: string[] = [];
    let length = 0;
    for (let index = 0; index < chunks.size; index++) {
      const chunk = chunks.get(index);
      if (chunk === undefined || (length += chunk.length) > MAX_BYTES) invalid();
      ordered.push(chunk);
    }
    encoded = ordered.join('');
  }
  if (encoded.length > MAX_BYTES || new TextEncoder().encode(encoded).byteLength > MAX_BYTES)
    invalid();
  let raw: unknown;
  try {
    raw = JSON.parse(encoded);
  } catch {
    invalid();
  }
  if (
    !object(raw) ||
    !keys(raw, ['schemaVersion', 'sources']) ||
    raw.schemaVersion !== 1 ||
    !Array.isArray(raw.sources) ||
    raw.sources.length > 100_000
  )
    invalid();
  const seen = new Set<string>();
  let count = 0;
  for (const source of raw.sources) {
    if (
      !object(source) ||
      !keys(source, ['scope', 'runId', 'token', 'absent', 'retained', 'texts']) ||
      typeof source.scope !== 'string' ||
      !UUID.test(source.scope) ||
      !id(source.runId) ||
      !token(source.token, source.scope) ||
      typeof source.absent !== 'boolean' ||
      !Array.isArray(source.retained) ||
      source.retained.length > 64 ||
      source.retained.some(
        (value) =>
          typeof value !== 'string' || value.length > 128 || (value !== '' && !isXmlNCName(value))
      ) ||
      new Set(source.retained).size !== source.retained.length ||
      !Array.isArray(source.texts) ||
      source.texts.length > 100_000
    )
      invalid();
    for (const nativeId of [
      source.runId,
      ...source.texts.map((text) => (object(text) ? text.id : null)),
    ]) {
      if (!id(nativeId) || ++count > 100_000) invalid();
      const key = JSON.stringify([source.scope, nativeId]);
      if (seen.has(key)) invalid();
      seen.add(key);
    }
    for (const text of source.texts)
      if (
        !object(text) ||
        !keys(text, ['id', 'token', 'space']) ||
        !id(text.id) ||
        !token(text.token, source.scope) ||
        !['absent', 'default', 'preserve'].includes(String(text.space))
      )
        invalid();
  }
  return raw.sources as unknown as readonly TrackedSourceBinding[];
}
export function withTrackedSourceBindings(
  node: OoxmlElement,
  sources: readonly TrackedSourceBinding[],
  preferredPrefix = 'sourceBindings'
): OoxmlElement {
  if (node.kind !== 'paragraph') invalid();
  const raw = JSON.stringify({ schemaVersion: 1, sources });
  if (new TextEncoder().encode(raw).byteLength > MAX_BYTES) invalid();
  const existing = node.attributes.find((attr) => attr.namespaceUri === TRACKED_SOURCE_BINDINGS_NS);
  const bindings = node.namespaceBindings;
  let prefix = existing?.prefix ?? preferredPrefix;
  for (
    let suffix = 1;
    bindings.some(
      (binding) => binding.prefix === prefix && binding.namespaceUri !== TRACKED_SOURCE_BINDINGS_NS
    );
    suffix++
  )
    prefix = preferredPrefix + suffix;
  const chunks: string[] = [];
  for (let start = 0; start < raw.length; ) {
    let end = Math.min(start + CHUNK_LENGTH, raw.length);
    const code = raw.charCodeAt(end - 1);
    if (end < raw.length && code >= 0xd800 && code <= 0xdbff) end--;
    chunks.push(raw.slice(start, end));
    start = end;
  }
  const next = {
    ...node,
    namespaceBindings:
      sources.length &&
      !bindings.some(
        (binding) =>
          binding.prefix === prefix && binding.namespaceUri === TRACKED_SOURCE_BINDINGS_NS
      )
        ? [...bindings, { prefix, namespaceUri: TRACKED_SOURCE_BINDINGS_NS }]
        : bindings,
    attributes: [
      ...node.attributes.filter((attr) => attr.namespaceUri !== TRACKED_SOURCE_BINDINGS_NS),
      ...(sources.length
        ? chunks.map((value, index) => ({
            kind: 'genericExtension' as const,
            namespaceUri: TRACKED_SOURCE_BINDINGS_NS,
            localName: chunks.length === 1 ? 'bindings' : 'bindings' + index,
            prefix,
            value,
          }))
        : []),
    ],
  } as OoxmlElement;
  if (sources.length) readTrackedSourceBindings(next);
  else {
    const referenced = new Set<string>();
    const aliases = next.namespaceBindings.filter(
      (binding) => binding.namespaceUri === TRACKED_SOURCE_BINDINGS_NS
    );
    const pending: OoxmlNode[] = [next];
    while (pending.length) {
      const current = pending.pop()!;
      if (current.kind === 'textValue') continue;
      for (const alias of aliases) {
        if (
          current.prefix === alias.prefix ||
          current.attributes.some(
            (attribute) =>
              attribute.prefix === alias.prefix || attribute.value.includes(alias.prefix + ':')
          )
        )
          referenced.add(alias.prefix);
      }
      for (const child of current.children) pending.push(child);
    }
    return {
      ...next,
      namespaceBindings: next.namespaceBindings.filter(
        (binding) =>
          binding.namespaceUri !== TRACKED_SOURCE_BINDINGS_NS || referenced.has(binding.prefix)
      ),
    } as OoxmlElement;
  }
  return next;
}
function stamp(
  node: OoxmlElement,
  source: TrackedSourceBinding,
  values: Readonly<Record<string, string>>,
  prefix: string
): OoxmlElement {
  const uri = ORIGIN_NS + '/' + source.scope;
  // Already scoped nodes retain their authored proof, including deferred resolution fields.
  if (node.attributes.some((attr) => attr.namespaceUri === uri)) return node;
  return {
    ...node,
    namespaceBindings: [...node.namespaceBindings, { prefix, namespaceUri: uri }],
    attributes: [
      ...node.attributes,
      ...Object.entries(values).map(([localName, value]) => ({
        kind: 'genericExtension' as const,
        namespaceUri: uri,
        localName,
        prefix,
        value,
      })),
    ],
  } as OoxmlElement;
}
/** Pure transfer. Foreign revisions, controls, and arbitrary run boundaries stay opaque. */
function paragraphBindings(node: OoxmlElement, scopes?: ReadonlySet<string>): OoxmlElement {
  let sources: readonly TrackedSourceBinding[];
  try {
    sources = readTrackedSourceBindings(node);
  } catch (error) {
    if (scopes) throw error;
    return node;
  }
  const selected = sources.filter((source) => !scopes || scopes.has(source.scope));
  if (!selected.length) return node;
  // Reserve declared prefixes and opaque QName-like attribute values throughout the
  // paragraph. Adding a local proof alias must not rebind an existing lexical value.
  const used = new Set<string>(),
    pending: OoxmlNode[] = [node];
  while (pending.length) {
    const current = pending.pop()!;
    if (current.kind === 'textValue') continue;
    if (current.prefix) used.add(current.prefix);
    for (const binding of current.namespaceBindings) used.add(binding.prefix);
    for (const attribute of current.attributes) {
      if (attribute.prefix) used.add(attribute.prefix);
      for (const match of attribute.value.matchAll(/([A-Za-z_][A-Za-z0-9_.-]*):/g))
        used.add(match[1]!);
    }
    for (const child of current.children) pending.push(child);
  }
  const prefixes = new Map<string, string>();
  for (const source of selected) {
    let prefix = 'runSource' + source.scope.replaceAll('-', '');
    while (used.has(prefix)) prefix += '_';
    if (!prefixes.has(source.scope)) {
      prefixes.set(source.scope, prefix);
      used.add(prefix);
    }
  }
  const byRun = new Map<string, TrackedSourceBinding[]>();
  for (const source of selected)
    byRun.set(source.runId, [...(byRun.get(source.runId) ?? []), source]);
  const visit = (current: OoxmlNode): OoxmlNode => {
    if (current.kind === 'paragraph' || current.kind === 'hyperlink') {
      const children = current.children.map(visit);
      return children.every((child, index) => child === current.children[index])
        ? current
        : ({ ...current, children } as OoxmlElement);
    }
    if (current.kind !== 'run') return current;
    let next: OoxmlElement = current;
    for (const source of byRun.get(current.id) ?? []) {
      next = stamp(
        next,
        source,
        {
          origin: source.token,
          identity: 'source',
          ...(source.absent ? { absent: '1' } : {}),
          ...(source.retained.length ? { retained: JSON.stringify(source.retained) } : {}),
        },
        prefixes.get(source.scope)!
      );
      next = {
        ...next,
        children: next.children.map((child) => {
          if (child.kind !== 'text' && child.kind !== 'deletedText') return child;
          const text = source.texts.find((text) => text.id === child.id);
          return text
            ? stamp(
                child,
                source,
                {
                  textOrigin: text.token,
                  space: text.space,
                  identity: 'source',
                },
                prefixes.get(source.scope)!
              )
            : child;
        }),
      } as OoxmlElement;
    }
    return next;
  };
  return withTrackedSourceBindings(
    visit(node) as OoxmlElement,
    sources.filter((source) => scopes && !scopes.has(source.scope))
  );
}
const exported = new WeakMap<OoxmlNode, OoxmlNode>();
/** Convert load-local bindings to self-contained proofs before native IDs change on XML reopening. */
export function materializeTrackedSourceBindings(
  root: OoxmlNode,
  scopes?: ReadonlySet<string>
): OoxmlNode {
  let count = 0;
  const visit = (node: OoxmlNode, depth: number): OoxmlNode => {
    if (++count > 1_000_000 || depth > 256) invalid();
    if (!scopes) {
      const cached = exported.get(node);
      if (cached) return cached;
    }
    let next = node;
    if (node.kind === 'paragraph') next = paragraphBindings(node, scopes);
    else if (node.kind !== 'textValue') {
      const children = node.children.map((child) => visit(child, depth + 1));
      if (!children.every((child, index) => child === node.children[index]))
        next = { ...node, children } as OoxmlElement;
    }
    if (!scopes) exported.set(node, next);
    return next;
  };
  return visit(root, 0);
}
