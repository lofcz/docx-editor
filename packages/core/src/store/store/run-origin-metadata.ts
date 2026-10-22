import { isMcPrefixListAttribute } from '../package/ooxml-serialize-namespaces.ts';
import { MAX_XML_DEPTH } from '../package/xml-reader.ts';
import { isXmlNCName } from '../package/qname.ts';
import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';

export const RUN_ORIGIN_NS = 'urn:docx-editor:tracked-run-origins:1';
export const RUN_ORIGIN_MAX_NODES = 100_000;
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
export const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
export const runProperties = (node: OoxmlNode) =>
  children(node).find((child) => child.kind === 'runProperties');
export const originNamespaceForScope = (scope: string) => `${RUN_ORIGIN_NS}/${scope}`;
export const isRunOriginNamespace = (uri: string) =>
  uri === RUN_ORIGIN_NS ||
  (uri.startsWith(`${RUN_ORIGIN_NS}/`) && UUID.test(uri.slice(RUN_ORIGIN_NS.length + 1)));
export const originNamespaces = (node: OoxmlNode | undefined): readonly string[] =>
  !node || node.kind === 'textValue'
    ? []
    : [...new Set(node.attributes.map((attr) => attr.namespaceUri).filter(isRunOriginNamespace))];
export const hasOriginMetadata = (node: OoxmlNode) => originNamespaces(node).length !== 0;

function fields(
  node: OoxmlNode,
  allowed: readonly string[],
  namespaceUri = RUN_ORIGIN_NS
): Map<string, string> | null {
  const result = new Map<string, string>();
  if (node.kind === 'textValue') return null;
  for (const attr of node.attributes) {
    if (attr.namespaceUri !== namespaceUri) continue;
    if (!allowed.includes(attr.localName) || result.has(attr.localName)) return null;
    result.set(attr.localName, attr.value);
  }
  return result;
}

export interface RunOrigin {
  readonly namespaceUri: string;
  readonly token: string;
  readonly identity: 'source' | null;
  readonly resolved: boolean;
  readonly scope: string;
  readonly absent: boolean;
  readonly retainedPrefixes: readonly string[];
}

function retainedPrefixes(raw: string | undefined): readonly string[] | null {
  if (raw === undefined) return [];
  if (raw.length > 8192) return null;
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return null;
  }
  if (
    !Array.isArray(value) ||
    value.length > 64 ||
    value.some(
      (item) => typeof item !== 'string' || item.length > 128 || (item !== '' && !isXmlNCName(item))
    )
  )
    return null;
  return new Set(value).size === value.length ? value : null;
}

export function encodeRetainedPrefixes(prefixes: readonly string[]): string | undefined {
  if (!prefixes.length) return undefined;
  const encoded = JSON.stringify(prefixes);
  if (!retainedPrefixes(encoded))
    throw new PackageEditRefusal(
      'Source namespace aliases exceed the provenance limit.',
      'resource-limit'
    );
  return encoded;
}

/** Unknown versions, extra fields, and malformed tokens remain opaque. */
export function readRunOrigins(node: OoxmlNode | undefined): readonly RunOrigin[] {
  return originNamespaces(node)
    .map((uri) => readRunOrigin(node, uri))
    .filter((origin): origin is RunOrigin => origin !== null);
}
export function readRunOrigin(
  node: OoxmlNode | undefined,
  namespaceUri?: string
): RunOrigin | null {
  if (namespaceUri === undefined) return readRunOrigins(node).at(-1) ?? null;
  if (!node || (node.kind !== 'runProperties' && node.kind !== 'run') || !hasOriginMetadata(node))
    return null;
  const data = fields(node, ['origin', 'absent', 'retained', 'identity', 'resolved'], namespaceUri);
  const token = data?.get('origin');
  if (!token || token.length !== 73) return null;
  const [scope, run] = token.split('/');
  if (!scope || !run || !UUID.test(scope) || !UUID.test(run)) return null;
  if (namespaceUri !== RUN_ORIGIN_NS && namespaceUri !== originNamespaceForScope(scope))
    return null;
  if (data!.has('absent') && data!.get('absent') !== '1') return null;
  if (data!.has('identity') && data!.get('identity') !== 'source') return null;
  if (
    data!.has('resolved') &&
    (data!.get('resolved') !== '1' || data!.get('identity') !== 'source')
  )
    return null;
  const retained = retainedPrefixes(data!.get('retained'));
  return retained
    ? {
        namespaceUri,
        token,
        identity: data!.get('identity') === 'source' ? 'source' : null,
        resolved: data!.get('resolved') === '1',
        scope,
        absent: data!.get('absent') === '1',
        retainedPrefixes: retained,
      }
    : null;
}

export interface RevisionOrigin {
  readonly namespaceUri: string;
  readonly scope: string;
  readonly retainedPrefixes: readonly string[];
}
export function readRevisionOrigins(node: OoxmlNode): readonly RevisionOrigin[] {
  return originNamespaces(node)
    .map((uri) => readRevisionOrigin(node, uri))
    .filter((origin): origin is RevisionOrigin => origin !== null);
}
export function readRevisionOrigin(node: OoxmlNode, namespaceUri?: string): RevisionOrigin | null {
  if (namespaceUri === undefined) return readRevisionOrigins(node).at(-1) ?? null;
  if (!hasOriginMetadata(node)) return null;
  const data = fields(node, ['scope', 'retained'], namespaceUri);
  const scope = data?.get('scope');
  const retained = data && retainedPrefixes(data.get('retained'));
  return scope &&
    UUID.test(scope) &&
    retained &&
    (namespaceUri === RUN_ORIGIN_NS || namespaceUri === originNamespaceForScope(scope))
    ? { namespaceUri, scope, retainedPrefixes: retained }
    : null;
}

export const readRevisionOriginScope = (node: OoxmlNode) => readRevisionOrigin(node)?.scope ?? null;

export interface TextOrigin {
  readonly token: string;
  readonly identity: 'source' | null;
  readonly resolved: boolean;
  readonly namespaceUri: string;
  readonly scope: string;
  readonly space: string;
}
export function readTextOrigins(node: OoxmlNode): readonly TextOrigin[] {
  return originNamespaces(node)
    .map((uri) => readTextOrigin(node, uri))
    .filter((origin): origin is TextOrigin => origin !== null);
}
export function readTextOrigin(node: OoxmlNode, namespaceUri?: string): TextOrigin | null {
  if (namespaceUri === undefined) return readTextOrigins(node).at(-1) ?? null;
  if (node.kind !== 'text' && node.kind !== 'deletedText') return null;
  const data = fields(node, ['textOrigin', 'space', 'identity', 'resolved'], namespaceUri);
  const token = data?.get('textOrigin');
  if (!token || token.length !== 73) return null;
  const [scope, text] = token.split('/');
  const space = data?.get('space');
  if (data?.has('resolved') && (data.get('resolved') !== '1' || data.get('identity') !== 'source'))
    return null;
  if (data?.has('identity') && data.get('identity') !== 'source') return null;
  return scope &&
    text &&
    UUID.test(scope) &&
    UUID.test(text) &&
    (namespaceUri === RUN_ORIGIN_NS || namespaceUri === originNamespaceForScope(scope)) &&
    (space === 'absent' || space === 'default' || space === 'preserve')
    ? {
        namespaceUri,
        scope,
        space,
        token,
        identity: data!.get('identity') === 'source' ? 'source' : null,
        resolved: data!.get('resolved') === '1',
      }
    : null;
}

/** Bound both work and recursion before building any replacement nodes. */
export function originNodes(root: OoxmlNode): readonly OoxmlNode[] {
  const result: OoxmlNode[] = [];
  const pending = [{ node: root, depth: 0 }];
  while (pending.length) {
    const { node, depth } = pending.pop()!;
    if (depth > MAX_XML_DEPTH + 1 || result.length >= RUN_ORIGIN_MAX_NODES)
      throw new PackageEditRefusal('Run-origin traversal exceeds its limit.', 'resource-limit');
    result.push(node);
    const nested = children(node);
    if (pending.length + nested.length > RUN_ORIGIN_MAX_NODES)
      throw new PackageEditRefusal('Run-origin traversal exceeds its limit.', 'resource-limit');
    for (let at = nested.length - 1; at >= 0; at--)
      pending.push({ node: nested[at]!, depth: depth + 1 });
  }
  return result;
}

export function mapOriginNodes(
  nodes: readonly OoxmlNode[],
  transform: (node: OoxmlNode, original: OoxmlNode) => OoxmlNode
): OoxmlNode {
  const mapped = new Map<OoxmlNode, OoxmlNode>();
  for (let at = nodes.length - 1; at >= 0; at--) {
    const node = nodes[at]!;
    let next = node;
    if (node.kind !== 'textValue') {
      const nested = node.children.map((child) => mapped.get(child) ?? child);
      if (nested.some((child, index) => child !== node.children[index]))
        next = { ...node, children: nested } as OoxmlNode;
    }
    mapped.set(node, transform(next, node));
  }
  return mapped.get(nodes[0]!)!;
}

/** Never shadow a prefix used anywhere in the annotated subtree. */
export function originPrefix(
  nodes: readonly OoxmlNode[],
  inherited: ReadonlyMap<string, string>,
  stem = 'runOrigin'
) {
  const used = new Set(inherited.keys());
  for (const node of nodes) {
    if (node.kind === 'textValue') continue;
    if (node.prefix) used.add(node.prefix);
    for (const attr of node.attributes) if (attr.prefix) used.add(attr.prefix);
    for (const binding of node.namespaceBindings) used.add(binding.prefix);
  }
  let prefix = stem;
  for (let suffix = 1; used.has(prefix); suffix++) prefix = `${stem}${suffix}`;
  return prefix;
}

export function withOriginMetadata(
  node: OoxmlElement,
  values: Readonly<Record<string, string>>,
  prefix: string
): OoxmlElement {
  const scope = values.scope ?? (values.origin ?? values.textOrigin ?? '').split('/')[0]!;
  if (!UUID.test(scope)) throw new PackageEditRefusal('Invalid provenance scope.', 'invalidArgs');
  const namespaceUri = originNamespaceForScope(scope);
  return {
    ...node,
    namespaceBindings: [...node.namespaceBindings, { prefix, namespaceUri }],
    attributes: [
      ...node.attributes,
      ...Object.entries(values).map(([localName, value]) => ({
        kind: 'genericExtension' as const,
        namespaceUri,
        prefix,
        localName,
        value,
      })),
    ],
  } as OoxmlElement;
}

export function withoutOriginMetadata(
  node: OoxmlElement,
  namespaceUri = RUN_ORIGIN_NS
): OoxmlElement {
  const retained = new Set(
    readRunOrigin(node, namespaceUri)?.retainedPrefixes ??
      readRevisionOrigin(node, namespaceUri)?.retainedPrefixes ??
      []
  );
  const candidates = node.namespaceBindings
    .filter((binding) => binding.namespaceUri === namespaceUri && !retained.has(binding.prefix))
    .map((binding) => binding.prefix);
  const used = new Set<string>();
  for (const child of originNodes(node)) {
    if (child.kind === 'textValue') continue;
    if (child.prefix) used.add(child.prefix);
    for (const attr of child.attributes) {
      if (child === node && attr.namespaceUri === namespaceUri) continue;
      if (attr.prefix) used.add(attr.prefix);
      const prefixes = isMcPrefixListAttribute(attr, child) ? attr.value.split(/\s+/) : [];
      for (const prefix of candidates)
        if (!used.has(prefix) && (attr.value.includes(`${prefix}:`) || prefixes.includes(prefix)))
          used.add(prefix);
    }
  }
  return {
    ...node,
    attributes: node.attributes.filter((attr) => attr.namespaceUri !== namespaceUri),
    namespaceBindings: node.namespaceBindings.filter(
      (binding) =>
        binding.namespaceUri !== namespaceUri ||
        binding.prefix === '' ||
        retained.has(binding.prefix) ||
        used.has(binding.prefix)
    ),
  } as OoxmlElement;
}
