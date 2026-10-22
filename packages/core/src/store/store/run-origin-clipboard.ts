import type { OoxmlNode } from '../package/ooxml-tree.ts';
import {
  children,
  hasOriginMetadata,
  mapOriginNodes,
  readRunOrigin,
  readTextOrigin,
  readRevisionOrigin,
  withoutOriginMetadata,
  RUN_ORIGIN_NS,
} from './run-origin-metadata.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';

/** A paste creates independent provenance; partial ordinary copies carry no pending scope. */
export function freshenCopiedRunOrigins(
  blocks: readonly OoxmlNode[],
  maxNodes: number,
  maxDepth: number
): readonly OoxmlNode[] {
  const groups: OoxmlNode[][] = [];
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
    groups.push(nodes);
  }
  if (!any) return blocks;
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
  const scopes = new Map<string, string>();
  for (const nodes of groups)
    for (const node of nodes) {
      const revision = readRevisionOrigin(node);
      const parent = parents.get(node);
      const valid =
        node.kind === 'revisionInsert' ||
        node.kind === 'revisionDelete' ||
        (node.kind !== 'textValue' &&
          node.localName === 'rPrChange' &&
          parent?.kind === 'runProperties' &&
          parents.get(parent)?.kind === 'run');
      if (revision && valid && ordinary(node) && !scopes.has(revision.scope))
        scopes.set(revision.scope, crypto.randomUUID());
    }
  const tokens = new Map<string, string>();
  return groups.map((nodes) =>
    mapOriginNodes(nodes, (node, original) => {
      if (node.kind === 'textValue' || !hasOriginMetadata(node)) return node;
      if (!ordinary(original)) return node;
      const parent = parents.get(original);
      const run = readRunOrigin(node);
      const candidate = readTextOrigin(node);
      const text = candidate && readRunOrigin(parent)?.scope === candidate.scope ? candidate : null;
      const revisionRole =
        node.kind === 'revisionInsert' ||
        node.kind === 'revisionDelete' ||
        (node.localName === 'rPrChange' &&
          parent?.kind === 'runProperties' &&
          parents.get(parent)?.kind === 'run');
      const revision = revisionRole ? readRevisionOrigin(node) : null;
      const oldScope = run?.scope ?? text?.scope ?? revision?.scope;
      if (!oldScope) return node;
      const scope = scopes.get(oldScope);
      if (!scope) return run || text ? withoutOriginMetadata(node) : node;
      return {
        ...node,
        attributes: node.attributes.map((attr) => {
          if (attr.namespaceUri !== RUN_ORIGIN_NS) return attr;
          if (attr.localName === 'scope') return { ...attr, value: scope };
          if (attr.localName !== 'origin' && attr.localName !== 'textOrigin') return attr;
          let token = tokens.get(attr.value);
          if (!token) {
            token = `${scope}/${crypto.randomUUID()}`;
            tokens.set(attr.value, token);
          }
          return { ...attr, value: token };
        }),
      } as OoxmlNode;
    })
  );
}
