import { parentNodeOf } from '../package/ooxml-edit.ts';
import { namespaceBindingsAt } from '../package/wml-namespace.ts';
import { isXmlNCName } from '../package/qname.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import {
  originNodes,
  readRunOrigin,
  runProperties,
  originNamespaceForScope,
} from './run-origin-metadata.ts';

/** Read source shells from surviving original markup, never from proposed characters. */
export function sourceRunShells(
  part: OoxmlPart,
  nodes: readonly OoxmlNode[],
  paragraphId: string,
  scope: string
): ReadonlyMap<string, OoxmlNode> {
  const shells = new Map<string, OoxmlNode>();
  for (const node of nodes) {
    if (node.kind !== 'run') continue;
    const origin = readRunOrigin(runProperties(node), originNamespaceForScope(scope));
    if (origin?.scope !== scope) continue;
    let owner = parentNodeOf(part, node.id);
    for (let depth = 0; owner && depth <= 256; depth++) {
      if (owner.id === paragraphId) {
        if (!shells.has(origin.token)) shells.set(origin.token, boundSourceShell(part, node));
        break;
      }
      if (owner.kind !== 'hyperlink' && owner.kind !== 'revisionDelete') break;
      owner = parentNodeOf(part, owner.id);
    }
  }
  return shells;
}

/** Inherit element metadata during authoring, while retaining all native identities and content. */
export function inheritSourceRunShell(node: OoxmlNode, source: OoxmlNode | undefined): OoxmlNode {
  if (node.kind !== 'run' || source?.kind !== 'run') return node;
  return {
    ...node,
    prefix: source.prefix,
    attributes: source.attributes,
    namespaceBindings: source.namespaceBindings,
  } as OoxmlNode;
}

/** Carry only referenced inherited aliases when a run moves into a new hyperlink context. */
function boundSourceShell(part: OoxmlPart, source: OoxmlNode): OoxmlNode {
  if (source.kind !== 'run') return source;
  const inherited = namespaceBindingsAt(part, source);
  const bindings = new Map(source.namespaceBindings.map((binding) => [binding.prefix, binding]));
  const needed = new Set<string>();
  for (const node of originNodes(source)) {
    if (node.kind === 'textValue') continue;
    if (node.prefix) needed.add(node.prefix);
    for (const attr of node.attributes) {
      if (attr.prefix) needed.add(attr.prefix);
      for (const token of attr.value.split(/\s+/)) {
        const pair = token.split(':');
        if (pair.length === 2 && pair.every(isXmlNCName)) needed.add(pair[0]!);
      }
    }
  }
  for (const prefix of needed) {
    const namespaceUri = inherited.get(prefix);
    if (namespaceUri && !bindings.has(prefix)) bindings.set(prefix, { prefix, namespaceUri });
  }
  return { ...source, namespaceBindings: [...bindings.values()] };
}
