import {
  contentControlLevelOf,
  findNode,
  parentNodeOf,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';

/** A complete contiguous range of native siblings. Descendants of selected roots occur once. */
export interface BlockSelectionRange {
  readonly success: true;
  readonly parent: OoxmlElement;
  readonly nodes: readonly OoxmlElement[];
}

/** Native block selection never treats runs, table rows, or property nodes as whole blocks. */
export function isSelectableBlock(node: OoxmlNode): node is OoxmlElement {
  return (
    node.kind === 'paragraph' ||
    node.kind === 'table' ||
    (node.kind === 'contentControl' && ['block', 'empty'].includes(contentControlLevelOf(node)))
  );
}

/**
 * Resolve explicit complete blocks in one story. The command never expands a partial or
 * disjoint request. Chrome supplies every sibling between the selection endpoints.
 */
export function resolveBlockSelectionRange(
  part: OoxmlPart,
  nodeIds: readonly string[]
): BlockSelectionRange | NativeEditFailure {
  const fail = (error: string): NativeEditFailure => ({
    success: false,
    code: 'invalid-input',
    error,
  });
  if (!nodeIds.length || nodeIds.length > 4096)
    return fail('Choose between 1 and 4096 complete blocks.');
  const unique = new Set(nodeIds);
  const nodes: OoxmlElement[] = [];
  for (const id of unique) {
    const node = findNode(part, id);
    if (!node || !isSelectableBlock(node))
      return fail('A selected block is missing or does not support whole-block editing.');
    let covered = false;
    for (let parent = parentNodeOf(part, id); parent; parent = parentNodeOf(part, parent.id)) {
      if (unique.has(parent.id)) {
        covered = true;
        break;
      }
    }
    if (!covered) nodes.push(node);
  }
  const parent = nodes[0] && parentNodeOf(part, nodes[0].id);
  if (!parent) return fail('The selected blocks have no owner.');
  if (nodes.some((node) => parentNodeOf(part, node.id)?.id !== parent.id))
    return fail('Choose complete blocks with the same native owner.');
  const selected = new Set(nodes.map((node) => node.id));
  const ordered = parent.children.filter((node) => selected.has(node.id)) as OoxmlElement[];
  const start = parent.children.findIndex((node) => node.id === ordered[0]!.id);
  if (parent.children.slice(start, start + ordered.length).some((node) => !selected.has(node.id)))
    return fail('Choose every complete block between the selection endpoints.');
  return { success: true, parent, nodes: ordered };
}
