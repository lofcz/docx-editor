import { secureRandomUuid } from '../package/random-uuid.ts';
import { materializeTrackedSourceBindings } from '../package/tracked-source-bindings.ts';
import type { OoxmlNode } from '../package/ooxml-tree.ts';
import {
  children,
  hasOriginMetadata,
  mapOriginNodes,
  readRunOrigin,
  readTextOrigin,
  readRevisionOrigin,
  withoutOriginMetadata,
  originNamespaces,
  originNamespaceForScope,
  readRunOrigins,
  readRevisionOrigins,
} from './run-origin-metadata.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';

/** A paste creates independent provenance; partial ordinary copies carry no pending scope. */
export function freshenCopiedRunOrigins(
  blocks: readonly OoxmlNode[],
  maxNodes: number,
  maxDepth: number
): readonly OoxmlNode[] {
  const groups: OoxmlNode[][] = [];
  const preparedBlocks: OoxmlNode[] = [];
  const parents = new Map<OoxmlNode, OoxmlNode>();
  let total = 0;
  let any = false;
  for (const block of blocks) {
    const nodes: OoxmlNode[] = [];
    const pending = [{ node: block, depth: 0 }];
    while (pending.length) {
      const { node, depth } = pending.pop()!;
      if (++total > maxNodes || depth > maxDepth)
        throw new PackageEditRefusal(
          'Copied provenance exceeds fragment limits.',
          'fragment-resource-budget'
        );
      nodes.push(node);
      any ||= hasOriginMetadata(node);
      for (const child of [...children(node)].reverse()) {
        parents.set(child, node);
        pending.push({ node: child, depth: depth + 1 });
      }
    }
    const projected = materializeTrackedSourceBindings(block);
    preparedBlocks.push(projected);
    if (projected === block) groups.push(nodes);
    else {
      const projectedNodes: OoxmlNode[] = [],
        pending = [projected];
      while (pending.length) {
        const node = pending.pop()!;
        projectedNodes.push(node);
        any ||= hasOriginMetadata(node);
        for (const child of [...children(node)].reverse()) {
          parents.set(child, node);
          pending.push(child);
        }
      }
      groups.push(projectedNodes);
    }
  }
  if (!any)
    return preparedBlocks.every((block, index) => block === blocks[index])
      ? blocks
      : preparedBlocks;
  const ordinary = (node: OoxmlNode): boolean => {
    let owner = parents.get(node);
    if (node.kind === 'text' || node.kind === 'deletedText' || node.kind === 'runProperties') {
      if (owner?.kind !== 'run') return false;
      owner = parents.get(owner);
    }
    if (node.kind !== 'textValue' && node.localName === 'rPrChange') {
      if (owner?.kind !== 'runProperties') return false;
      owner = parents.get(owner);
      if (owner?.kind !== 'run') return false;
      owner = parents.get(owner);
    }
    for (let depth = 0; owner && depth <= maxDepth; depth++) {
      if (owner.kind === 'paragraph') return true;
      if (
        owner.kind !== 'hyperlink' &&
        owner.kind !== 'revisionInsert' &&
        owner.kind !== 'revisionDelete'
      )
        return false;
      owner = parents.get(owner);
    }
    return false;
  };
  const used = new Set<string>();
  for (const nodes of groups)
    for (const node of nodes) {
      if (node.kind === 'textValue') continue;
      if (node.prefix) used.add(node.prefix);
      for (const binding of node.namespaceBindings) used.add(binding.prefix);
      for (const attr of node.attributes) if (attr.prefix) used.add(attr.prefix);
    }
  let suffix = 0;
  const mintPrefix = () => {
    let prefix = 'runOriginCopy';
    while (used.has(prefix)) prefix = `runOriginCopy${++suffix}`;
    used.add(prefix);
    return prefix;
  };
  const key = (namespaceUri: string, scope: string) => `${namespaceUri}|${scope}`;
  const scopes = new Map<string, { scope: string; namespaceUri: string; prefix: string }>();
  const revisionRole = (node: OoxmlNode, original: OoxmlNode) => {
    const parent = parents.get(original);
    return (
      node.kind === 'revisionInsert' ||
      node.kind === 'revisionDelete' ||
      (node.kind !== 'textValue' &&
        node.localName === 'rPrChange' &&
        parent?.kind === 'runProperties' &&
        parents.get(parent)?.kind === 'run')
    );
  };
  for (const nodes of groups)
    for (const node of nodes) {
      if (!revisionRole(node, node) || !ordinary(node)) continue;
      for (const origin of readRevisionOrigins(node)) {
        const id = key(origin.namespaceUri, origin.scope);
        if (!scopes.has(id)) {
          const scope = secureRandomUuid();
          scopes.set(id, {
            scope,
            namespaceUri: originNamespaceForScope(scope),
            prefix: mintPrefix(),
          });
        }
      }
    }
  // A completed peer proof can remain necessary for a copied pending proposal.
  const byScope = new Map<string, Set<string>>(),
    bySource = new Map<string, Set<string>>();
  for (const nodes of groups)
    for (const node of nodes) {
      if (!ordinary(node)) continue;
      for (const origin of readRunOrigins(node)) {
        const id = key(origin.namespaceUri, origin.scope);
        if (origin.identity !== 'source' || (!origin.resolved && !scopes.has(id))) continue;
        const source = origin.token.slice(37);
        const sources = byScope.get(id) ?? new Set<string>();
        sources.add(source);
        byScope.set(id, sources);
        const owners = bySource.get(source) ?? new Set<string>();
        owners.add(id);
        bySource.set(source, owners);
      }
    }
  const pending = [...scopes.keys()],
    visited = new Set<string>();
  while (pending.length) {
    const id = pending.pop()!;
    if (visited.has(id)) continue;
    visited.add(id);
    for (const source of byScope.get(id) ?? [])
      for (const other of bySource.get(source) ?? []) {
        if (!scopes.has(other)) {
          const scope = secureRandomUuid();
          scopes.set(other, {
            scope,
            namespaceUri: originNamespaceForScope(scope),
            prefix: mintPrefix(),
          });
        }
        if (!visited.has(other)) pending.push(other);
      }
  }
  const tokens = new Map<string, string>();
  return groups.map((nodes) =>
    mapOriginNodes(nodes, (node, original) => {
      if (node.kind === 'textValue' || !hasOriginMetadata(node) || !ordinary(original)) return node;
      let next = node;
      const parent = parents.get(original);
      for (const namespaceUri of originNamespaces(original)) {
        const run = readRunOrigin(original, namespaceUri),
          candidate = readTextOrigin(original, namespaceUri);
        const text =
          candidate &&
          readRunOrigins(parent).some(
            (origin) => origin.namespaceUri === namespaceUri && origin.scope === candidate.scope
          )
            ? candidate
            : null;
        const revision = revisionRole(node, original)
          ? readRevisionOrigin(original, namespaceUri)
          : null;
        const oldScope = run?.scope ?? text?.scope ?? revision?.scope;
        if (!oldScope) continue;
        const target = scopes.get(key(namespaceUri, oldScope));
        if (!target) {
          if (run || text) next = withoutOriginMetadata(next, namespaceUri);
          continue;
        }
        next = {
          ...next,
          namespaceBindings: [
            ...next.namespaceBindings,
            { prefix: target.prefix, namespaceUri: target.namespaceUri },
          ],
          attributes: next.attributes.map((attr) => {
            if (attr.namespaceUri !== namespaceUri) return attr;
            let value = attr.value;
            if (attr.localName === 'scope') value = target.scope;
            if (attr.localName === 'origin' || attr.localName === 'textOrigin') {
              const identity = run?.identity ?? text?.identity;
              const id =
                identity === 'source'
                  ? `${attr.localName}|source|${attr.value.slice(37)}`
                  : `${namespaceUri}|${attr.value}`;
              let source = tokens.get(id);
              if (!source) {
                source = secureRandomUuid();
                tokens.set(id, source);
              }
              value = `${target.scope}/${source}`;
            }
            return { ...attr, namespaceUri: target.namespaceUri, prefix: target.prefix, value };
          }),
        } as typeof node;
      }
      return next;
    })
  );
}
