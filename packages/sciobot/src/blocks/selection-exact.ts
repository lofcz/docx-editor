import {
  findNode,
  parentNodeOf,
  type OoxmlElement,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { isSelectableBlock } from './selection-range';

export interface ExactBlockSelection {
  readonly success: true;
  /** Selected roots in document order. An enclosing selected root owns its descendants. */
  readonly nodes: readonly OoxmlElement[];
  /** Each original owner and its exact selected children, without filling selection gaps. */
  readonly owners: readonly {
    readonly parent: OoxmlElement;
    readonly nodes: readonly OoxmlElement[];
  }[];
}

/** Resolve explicit whole blocks without converting disjoint selections into text ranges. */
export function resolveExactBlockSelection(
  part: OoxmlPart,
  nodeIds: readonly string[]
): ExactBlockSelection | NativeEditFailure {
  const fail = (error: string): NativeEditFailure => ({
    success: false,
    code: 'invalid-input',
    error,
  });
  if (!nodeIds.length || nodeIds.length > 4096)
    return fail('Choose between 1 and 4096 complete blocks.');
  const selected = new Set(nodeIds);
  const roots: { node: OoxmlElement; parent: OoxmlElement; path: number[] }[] = [];
  const positions = new Map<string, ReadonlyMap<string, number>>();
  let visited = 0;
  for (const id of selected) {
    const node = findNode(part, id),
      parent = parentNodeOf(part, id);
    if (!node || !isSelectableBlock(node) || !parent)
      return fail('A selected block is missing or does not support whole-block editing.');
    const path: number[] = [];
    let child = node,
      owner: OoxmlElement | null = parent,
      covered = false;
    while (owner) {
      if (path.length >= 64) return fail('The selected block exceeds the supported nesting limit.');
      if (selected.has(owner.id)) covered = true;
      let indices = positions.get(owner.id);
      if (!indices) {
        visited += owner.children.length;
        if (visited > 500_000) return fail('The selected owners exceed the supported node limit.');
        indices = new Map(owner.children.map((entry, index) => [entry.id, index]));
        positions.set(owner.id, indices);
      }
      const index = indices.get(child.id);
      if (index === undefined) return fail('A selected block has an invalid native owner.');
      path.push(index);
      child = owner;
      owner = parentNodeOf(part, owner.id);
    }
    if (!covered) roots.push({ node, parent, path: path.reverse() });
  }
  roots.sort((a, b) => {
    for (let i = 0; i < Math.min(a.path.length, b.path.length); i++)
      if (a.path[i] !== b.path[i]) return a.path[i]! - b.path[i]!;
    return a.path.length - b.path.length;
  });
  const owners = new Map<string, { parent: OoxmlElement; nodes: OoxmlElement[] }>();
  for (const root of roots) {
    let owner = owners.get(root.parent.id);
    if (!owner) {
      owner = { parent: root.parent, nodes: [] };
      owners.set(root.parent.id, owner);
    }
    owner.nodes.push(root.node);
  }
  return { success: true, nodes: roots.map((root) => root.node), owners: [...owners.values()] };
}
