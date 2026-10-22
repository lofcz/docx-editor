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
export const hasOriginMetadata = (node: OoxmlNode) =>
  node.kind !== 'textValue' && node.attributes.some((attr) => attr.namespaceUri === RUN_ORIGIN_NS);

function fields(node: OoxmlNode, allowed: readonly string[]): Map<string, string> | null {
  const result = new Map<string, string>();
  if (node.kind === 'textValue') return null;
  for (const attr of node.attributes) {
    if (attr.namespaceUri !== RUN_ORIGIN_NS) continue;
    if (!allowed.includes(attr.localName) || result.has(attr.localName)) return null;
    result.set(attr.localName, attr.value);
  }
  return result;
}

export interface RunOrigin {
  readonly token: string;
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
export function readRunOrigin(node: OoxmlNode | undefined): RunOrigin | null {
  if (!node || (node.kind !== 'runProperties' && node.kind !== 'run') || !hasOriginMetadata(node))
    return null;
  const data = fields(node, ['origin', 'absent', 'retained']);
  const token = data?.get('origin');
  if (!token || token.length !== 73) return null;
  const [scope, run] = token.split('/');
  if (!scope || !run || !UUID.test(scope) || !UUID.test(run)) return null;
  if (data!.has('absent') && data!.get('absent') !== '1') return null;
  const retained = retainedPrefixes(data!.get('retained'));
  return retained
    ? {
        token,
        scope,
        absent: data!.get('absent') === '1',
        retainedPrefixes: retained,
      }
    : null;
}

export function readRevisionOrigin(
  node: OoxmlNode
): { scope: string; retainedPrefixes: readonly string[] } | null {
  if (!hasOriginMetadata(node)) return null;
  const data = fields(node, ['scope', 'retained']);
  const scope = data?.get('scope');
  const retained = data && retainedPrefixes(data.get('retained'));
  return scope && UUID.test(scope) && retained ? { scope, retainedPrefixes: retained } : null;
}

export const readRevisionOriginScope = (node: OoxmlNode) => readRevisionOrigin(node)?.scope ?? null;

export function readTextOrigin(node: OoxmlNode): { scope: string; space: string } | null {
  if (node.kind !== 'text' && node.kind !== 'deletedText') return null;
  const data = fields(node, ['textOrigin', 'space']);
  const token = data?.get('textOrigin');
  if (!token || token.length !== 73) return null;
  const [scope, text] = token.split('/');
  const space = data?.get('space');
  return scope &&
    text &&
    UUID.test(scope) &&
    UUID.test(text) &&
    (space === 'absent' || space === 'default' || space === 'preserve')
    ? { scope, space }
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
export function originPrefix(nodes: readonly OoxmlNode[], inherited: ReadonlyMap<string, string>) {
  const used = new Set(inherited.keys());
  for (const node of nodes) {
    if (node.kind === 'textValue') continue;
    if (node.prefix) used.add(node.prefix);
    for (const attr of node.attributes) if (attr.prefix) used.add(attr.prefix);
    for (const binding of node.namespaceBindings) used.add(binding.prefix);
  }
  let prefix = 'runOrigin';
  for (let suffix = 1; used.has(prefix); suffix++) prefix = `runOrigin${suffix}`;
  return prefix;
}

export function withOriginMetadata(
  node: OoxmlElement,
  values: Readonly<Record<string, string>>,
  prefix: string
): OoxmlElement {
  return {
    ...node,
    namespaceBindings: [...node.namespaceBindings, { prefix, namespaceUri: RUN_ORIGIN_NS }],
    attributes: [
      ...node.attributes,
      ...Object.entries(values).map(([localName, value]) => ({
        kind: 'genericExtension' as const,
        namespaceUri: RUN_ORIGIN_NS,
        prefix,
        localName,
        value,
      })),
    ],
  } as OoxmlElement;
}

export function withoutOriginMetadata(node: OoxmlElement): OoxmlElement {
  const retained = new Set(
    readRunOrigin(node)?.retainedPrefixes ?? readRevisionOrigin(node)?.retainedPrefixes ?? []
  );
  const used = new Set(
    originNodes(node)
      .slice(1)
      .flatMap((child) =>
        child.kind === 'textValue'
          ? []
          : child.attributes
              .filter((attr) => attr.namespaceUri === RUN_ORIGIN_NS)
              .map((attr) => attr.prefix)
      )
  );
  return {
    ...node,
    attributes: node.attributes.filter((attr) => attr.namespaceUri !== RUN_ORIGIN_NS),
    namespaceBindings: node.namespaceBindings.filter(
      (binding) =>
        binding.namespaceUri !== RUN_ORIGIN_NS ||
        retained.has(binding.prefix) ||
        used.has(binding.prefix)
    ),
  } as OoxmlElement;
}
