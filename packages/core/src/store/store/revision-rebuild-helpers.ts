import { isMcPrefixListAttribute } from '../package/ooxml-serialize-namespaces.ts';
import { MAX_XML_DEPTH, XML_DEFAULT_MAX_ELEMENTS } from '../package/xml-reader.ts';
import { PackageEditRefusal, isPackageEditRefusal } from './package-edit-refusal.ts';
import type { TreeOpRejection } from './tree-op-types.ts';
// Node-level helpers for rebuilding a part after a tracked-change decision.

import { parentNodeOf } from '../package/ooxml-edit.ts';
import { declaredLockOf } from './tree-op-nodes.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';

export type Resolution = 'unwrap' | 'remove' | 'restore';

/**
 * The controls a removal must keep: a control locked against removal (`sdtLocked`,
 * `sdtContentLocked`) stays, emptied, where the removed content held it. The decision then
 * restores its placeholder.
 */
export function lockedControlShells(node: OoxmlNode): OoxmlNode[] {
  if (node.kind === 'textValue') return [];
  if (node.kind === 'contentControl' && declaredLockOf(node).wrapper)
    return [
      {
        ...node,
        children: node.children.map((child) =>
          child.kind === 'contentControlContent' ? { ...child, children: [] } : child
        ),
      } as OoxmlNode,
    ];
  return node.children.flatMap(lockedControlShells);
}

/** How one wrapper resolves, given the action. */
export function resolutionOf(node: OoxmlElement, action: 'accept' | 'reject'): Resolution {
  switch (node.kind) {
    case 'revisionInsert':
    case 'revisionMoveTo':
      // Accepting keeps the content and drops the wrapper; rejecting takes both.
      return action === 'accept' ? 'unwrap' : 'remove';
    case 'revisionDelete':
    case 'revisionMoveFrom':
      // Accepting removes the content; rejecting restores it, which also turns `w:delText`
      // back into `w:t`.
      return action === 'accept' ? 'remove' : 'restore';
    default:
      return 'remove';
  }
}

function textElementFrom(node: OoxmlElement): OoxmlElement {
  return { ...node, kind: 'text', localName: 't' } as OoxmlElement;
}

/**
 * Undo the deleted forms of run content when a deletion is rejected.
 *
 * `w:delText` is the deleted form of `w:t`, and `w:delInstrText` is the deleted form of
 * `w:instrText` (`EG_RunInnerContent`, §17.3.3.7 / §17.16.23). BOTH have to be inverted.
 * Restoring the text but leaving the instruction in its deleted form outside any `w:del` means
 * Word stops reading it as an instruction: the field's code is silently lost and the field
 * never updates again.
 */
export function withDeletedTextRestored(node: OoxmlNode): OoxmlNode {
  if (node.kind === 'textValue') return node;
  const children = node.children.map(withDeletedTextRestored);
  const rebuilt = children.some((child, index) => child !== node.children[index])
    ? ({ ...node, children } as OoxmlElement)
    : node;
  if (rebuilt.kind === 'deletedText') return textElementFrom(rebuilt);
  // `w:delInstrText` has no typed kind — it is layout-inert either way — so it is matched by
  // name and renamed in place.
  if (
    rebuilt.kind === 'generic' &&
    rebuilt.namespaceUri === WML_NAMESPACE_URI &&
    rebuilt.localName === 'delInstrText'
  ) {
    return { ...rebuilt, localName: 'instrText' } as OoxmlElement;
  }
  return rebuilt;
}

/**
 * Every node on the path from the part root to one of `targets`, the targets included.
 *
 * A rebuild descends only into these. Every other subtree holds nothing the decision changes,
 * so it is kept as it is, and the rebuild costs the paths it changes instead of the part.
 */
export function rebuildReach(part: OoxmlPart, targets: Iterable<string>): Set<string> {
  const reach = new Set<string>();
  for (const target of targets) {
    let id: string | undefined = target;
    while (id !== undefined && !reach.has(id)) {
      reach.add(id);
      id = parentNodeOf(part, id)?.id;
    }
  }
  return reach;
}

/** Keep aliases the surviving content needs when its revision wrapper disappears. */
export function withRevisionParentBindings(node: OoxmlNode, parent: OoxmlElement): OoxmlNode {
  if (node.kind === 'textValue' || !parent.namespaceBindings.length) return node;
  const own = new Set(node.namespaceBindings.map((binding) => binding.prefix));
  const inherited = parent.namespaceBindings.filter((binding) => !own.has(binding.prefix));
  if (!inherited.length) return node;
  const needed = new Set(
    inherited.filter((binding) => binding.prefix === '').map((binding) => binding.prefix)
  );
  const pending: { node: OoxmlNode; candidates: ReadonlySet<string>; depth: number }[] = [
    { node, candidates: new Set(inherited.map((binding) => binding.prefix)), depth: 0 },
  ];
  let visited = 0;
  while (pending.length && needed.size < inherited.length) {
    const frame = pending.pop()!;
    if (++visited > XML_DEFAULT_MAX_ELEMENTS * 2 + 1 || frame.depth > MAX_XML_DEPTH + 1)
      throw new PackageEditRefusal(
        'Revision namespace traversal exceeds its limit.',
        'resource-limit'
      );
    const current = frame.node;
    if (current.kind === 'textValue') continue;
    const candidates = current.namespaceBindings.length
      ? new Set(
          [...frame.candidates].filter(
            (prefix) => !current.namespaceBindings.some((binding) => binding.prefix === prefix)
          )
        )
      : frame.candidates;
    if (!candidates.size) continue;
    if (current.prefix && candidates.has(current.prefix)) needed.add(current.prefix);
    for (const attr of current.attributes) {
      if (attr.prefix && candidates.has(attr.prefix)) needed.add(attr.prefix);
      const prefixes = isMcPrefixListAttribute(attr, current) ? attr.value.split(/\s+/) : [];
      for (const prefix of candidates)
        if (!needed.has(prefix) && (attr.value.includes(`${prefix}:`) || prefixes.includes(prefix)))
          needed.add(prefix);
    }
    for (const child of current.children)
      pending.push({ node: child, candidates, depth: frame.depth + 1 });
  }
  if (!needed.size) return node;
  return {
    ...node,
    namespaceBindings: [
      ...inherited.filter((binding) => needed.has(binding.prefix)),
      ...node.namespaceBindings,
    ],
  } as OoxmlNode;
}

/** Native revision decisions report bounded rebuild refusals through their result contract. */
export function tryRevisionRebuild(
  run: () => readonly OoxmlNode[]
):
  | { readonly ok: true; readonly nodes: readonly OoxmlNode[] }
  | { readonly ok: false; readonly reason: TreeOpRejection } {
  try {
    return { ok: true, nodes: run() };
  } catch (error) {
    if (isPackageEditRefusal(error)) return { ok: false, reason: error.reason };
    throw error;
  }
}
