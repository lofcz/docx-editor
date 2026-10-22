import type { OoxmlNode } from '../package/ooxml-tree.ts';
import {
  children,
  mapOriginNodes,
  originNodes,
  readRunOrigins,
  readTextOrigins,
  readRevisionOrigins,
} from './run-origin-metadata.ts';

/** Keep a consumed proof until concurrent proposals from the same source have all settled. */
export function settleDeferredRunOrigins(
  root: OoxmlNode,
  selected: ReadonlySet<string>,
  hints: readonly { readonly scope: string; readonly source: string }[] = []
): {
  readonly root: OoxmlNode;
  readonly scopes: ReadonlySet<string>;
} {
  const nodes = originNodes(root);
  const live = new Set(
    nodes
      .filter(
        (node) =>
          node.kind !== 'textValue' &&
          (node.kind.startsWith('revision') || /PrChange$/.test(node.localName))
      )
      .flatMap(readRevisionOrigins)
      .map((origin) => origin.scope)
  );
  const clean = new Set([...selected].filter((scope) => !live.has(scope)));
  const origins = nodes
    .flatMap(readRunOrigins)
    .filter(
      (origin) =>
        origin.identity === 'source' &&
        (clean.has(origin.scope) || live.has(origin.scope) || origin.resolved)
    );
  const bySource = new Map<string, Set<string>>(),
    byScope = new Map<string, Set<string>>();
  for (const origin of origins) {
    const id = origin.token.slice(37);
    const scopes = bySource.get(id) ?? new Set<string>();
    scopes.add(origin.scope);
    bySource.set(id, scopes);
    const sources = byScope.get(origin.scope) ?? new Set<string>();
    sources.add(id);
    byScope.set(origin.scope, sources);
  }
  // A rejected wrapper can remove the last current node carrying the selected proof.
  for (const hint of hints) {
    if (!clean.has(hint.scope)) continue;
    const scopes = bySource.get(hint.source) ?? new Set<string>();
    scopes.add(hint.scope);
    bySource.set(hint.source, scopes);
    const sources = byScope.get(hint.scope) ?? new Set<string>();
    sources.add(hint.source);
    byScope.set(hint.scope, sources);
  }
  const visited = new Set<string>(),
    deferred = new Set<string>();
  for (const scope of [...clean]) {
    if (visited.has(scope) || !byScope.has(scope)) continue;
    const component = new Set<string>(),
      pending = [scope];
    while (pending.length) {
      const current = pending.pop()!;
      if (component.has(current)) continue;
      component.add(current);
      visited.add(current);
      for (const source of byScope.get(current) ?? [])
        for (const other of bySource.get(source) ?? [])
          if (!component.has(other)) pending.push(other);
    }
    if ([...component].some((item) => live.has(item))) {
      for (const item of component) if (clean.delete(item)) deferred.add(item);
    } else for (const item of component) clean.add(item);
  }
  if (!deferred.size) return { root, scopes: clean };
  const ordinary = new Set([root.id]);
  for (const node of nodes) {
    if (!ordinary.has(node.id)) continue;
    for (const child of children(node))
      if (
        ((node.kind === 'paragraph' || node.kind === 'hyperlink') &&
          (child.kind === 'run' || child.kind === 'hyperlink')) ||
        (node.kind === 'run' && (child.kind === 'text' || child.kind === 'deletedText'))
      )
        ordinary.add(child.id);
  }
  const marked = mapOriginNodes(nodes, (node) => {
    if (!ordinary.has(node.id) || node.kind === 'textValue') return node;
    const known = [...readRunOrigins(node), ...readTextOrigins(node)].filter(
      (origin) => origin.identity === 'source' && deferred.has(origin.scope) && !origin.resolved
    );
    if (!known.length) return node;
    return {
      ...node,
      attributes: [
        ...node.attributes,
        ...known.map((origin) => ({
          kind: 'genericExtension' as const,
          namespaceUri: origin.namespaceUri,
          prefix: node.attributes.find((attr) => attr.namespaceUri === origin.namespaceUri)!.prefix,
          localName: 'resolved',
          value: '1',
        })),
      ],
    } as OoxmlNode;
  });
  return { root: marked, scopes: clean };
}
