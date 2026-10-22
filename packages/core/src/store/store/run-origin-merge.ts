import { canonicalOoxmlFingerprint, type OoxmlNode } from '../package/ooxml-tree.ts';
import { sameInlineTextMetadata } from './tree-op-inline-slices.ts';
import {
  children,
  readRunOrigins,
  readTextOrigins,
  runProperties,
  withoutOriginMetadata,
  type RunOrigin,
  type TextOrigin,
} from './run-origin-metadata.ts';

type Origin = RunOrigin | TextOrigin;
function sameSource(a: Origin, b: Origin): boolean {
  return (
    (a.token === b.token && a.namespaceUri === b.namespaceUri) ||
    (a.identity === 'source' && b.identity === 'source' && a.token.slice(37) === b.token.slice(37))
  );
}
function selectedOrigins(node: OoxmlNode, scopes: ReadonlySet<string>): readonly Origin[] {
  return [...readRunOrigins(node), ...readTextOrigins(node)].filter((origin) =>
    scopes.has(origin.scope)
  );
}
function withoutSelected(node: OoxmlNode, scopes: ReadonlySet<string>): OoxmlNode {
  if (node.kind === 'textValue') return node;
  let next = node;
  for (const origin of selectedOrigins(node, scopes))
    next = withoutOriginMetadata(next, origin.namespaceUri);
  return next;
}
function attributes(node: OoxmlNode): string {
  return node.kind === 'textValue'
    ? ''
    : JSON.stringify(
        node.attributes.map((attr) => [attr.namespaceUri, attr.localName, attr.value]).sort()
      );
}
/** Keep both selected proofs until the normal resolution cleanup consumes their aliases. */
function combinedMetadata(left: OoxmlNode, right: OoxmlNode, scopes: ReadonlySet<string>) {
  if (left.kind === 'textValue' || right.kind === 'textValue') return null;
  const namespaces = new Set(selectedOrigins(right, scopes).map((origin) => origin.namespaceUri));
  const present = new Set(left.attributes.map((attr) => `${attr.namespaceUri}|${attr.localName}`));
  const attributes = [
    ...left.attributes,
    ...right.attributes.filter(
      (attr) =>
        namespaces.has(attr.namespaceUri) && !present.has(`${attr.namespaceUri}|${attr.localName}`)
    ),
  ];
  const bindings = new Map(left.namespaceBindings.map((binding) => [binding.prefix, binding]));
  for (const binding of right.namespaceBindings) {
    const prior = bindings.get(binding.prefix);
    if (prior && prior.namespaceUri !== binding.namespaceUri) return null;
    if (!prior) bindings.set(binding.prefix, binding);
  }
  return { attributes, namespaceBindings: [...bindings.values()] };
}

function joinOriginText(
  left: OoxmlNode,
  right: OoxmlNode,
  scopes: ReadonlySet<string>
): OoxmlNode | null {
  if (
    left.kind !== 'text' ||
    right.kind !== 'text' ||
    left.children.length !== 1 ||
    right.children.length !== 1 ||
    left.children[0]?.kind !== 'textValue' ||
    right.children[0]?.kind !== 'textValue'
  )
    return null;
  const a = readTextOrigins(left).filter((origin) => scopes.has(origin.scope));
  const b = readTextOrigins(right).filter((origin) => scopes.has(origin.scope));
  if (
    !a.some((origin) =>
      b.some((other) => sameSource(origin, other) && origin.space === other.space)
    )
  )
    return null;
  if (!sameInlineTextMetadata(withoutSelected(left, scopes), withoutSelected(right, scopes)))
    return null;
  const metadata = combinedMetadata(left, right, scopes);
  return metadata
    ? ({
        ...left,
        ...metadata,
        children: [
          { ...left.children[0], value: left.children[0].value + right.children[0].value },
        ],
      } as OoxmlNode)
    : null;
}

/** Join only provenance-proven fragments with identical native shells and current properties. */
export function joinOriginRuns(
  left: OoxmlNode,
  right: OoxmlNode,
  scopes: ReadonlySet<string>
): OoxmlNode | null {
  if (left.kind !== 'run' || right.kind !== 'run') return null;
  const lp = runProperties(left),
    rp = runProperties(right);
  const a = readRunOrigins(left).filter((origin) => scopes.has(origin.scope));
  const b = readRunOrigins(right).filter((origin) => scopes.has(origin.scope));
  if (
    !a.some((origin) =>
      b.some(
        (other) =>
          sameSource(origin, other) &&
          origin.absent === other.absent &&
          origin.retainedPrefixes.join('\0') === other.retainedPrefixes.join('\0')
      )
    )
  )
    return null;
  if (attributes(withoutSelected(left, scopes)) !== attributes(withoutSelected(right, scopes)))
    return null;
  if ((lp ? canonicalOoxmlFingerprint(lp) : null) !== (rp ? canonicalOoxmlFingerprint(rp) : null))
    return null;
  const metadata = combinedMetadata(left, right, scopes);
  if (!metadata) return null;
  const content: OoxmlNode[] = left.children.filter((child) => child !== lp);
  for (const node of children(right).filter((child) => child !== rp)) {
    const previous = content.at(-1),
      merged = previous && joinOriginText(previous, node, scopes);
    if (merged) content[content.length - 1] = merged;
    else content.push(node);
  }
  return { ...left, ...metadata, children: [...(lp ? [lp] : []), ...content] } as OoxmlNode;
}
