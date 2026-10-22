import { canonicalOoxmlFingerprint, type OoxmlNode } from '../package/ooxml-tree.ts';
import { sameInlineTextMetadata } from './tree-op-inline-slices.ts';
import { children, readRunOrigin, runProperties, RUN_ORIGIN_NS } from './run-origin-metadata.ts';

function attributes(node: OoxmlNode): string {
  return node.kind === 'textValue'
    ? ''
    : JSON.stringify(
        node.attributes
          .filter((attr) => attr.namespaceUri !== RUN_ORIGIN_NS)
          .map((attr) => [attr.namespaceUri, attr.localName, attr.value])
          .sort()
      );
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
  const a = readRunOrigin(left),
    b = readRunOrigin(right);
  if (
    !a ||
    !b ||
    a.token !== b.token ||
    a.absent !== b.absent ||
    a.retainedPrefixes.join('\0') !== b.retainedPrefixes.join('\0') ||
    !scopes.has(a.scope) ||
    attributes(left) !== attributes(right)
  )
    return null;
  if ((lp ? canonicalOoxmlFingerprint(lp) : null) !== (rp ? canonicalOoxmlFingerprint(rp) : null))
    return null;
  const bindings = new Map(left.namespaceBindings.map((binding) => [binding.prefix, binding]));
  for (const binding of right.namespaceBindings) {
    const prior = bindings.get(binding.prefix);
    if (prior && prior.namespaceUri !== binding.namespaceUri) return null;
    if (!prior) bindings.set(binding.prefix, binding);
  }
  const content: OoxmlNode[] = left.children.filter((child) => child !== lp);
  for (const node of children(right).filter((child) => child !== rp)) {
    const previous = content.at(-1);
    if (
      previous?.kind === 'text' &&
      node.kind === 'text' &&
      sameInlineTextMetadata(previous, node) &&
      previous.children.length === 1 &&
      node.children.length === 1 &&
      previous.children[0]?.kind === 'textValue' &&
      node.children[0]?.kind === 'textValue'
    ) {
      content[content.length - 1] = {
        ...previous,
        children: [
          { ...previous.children[0], value: previous.children[0].value + node.children[0].value },
        ],
      } as OoxmlNode;
    } else content.push(node);
  }
  return {
    ...left,
    namespaceBindings: [...bindings.values()],
    children: [...(lp ? [lp] : []), ...content],
  } as OoxmlNode;
}
