import type { BlockInset } from '@docx-editor.dev/core/editor';
import {
  blockStoryContainerChildren,
  contentControlContentNodeOf,
  findNode,
  MAX_STORY_SDT_NESTING,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { BlockIndex } from '../blocks/block-index';

/** Padding a container view gives its native content, in CSS pixels at 100% zoom. */
export interface ReactBlockInset {
  readonly left?: number;
  readonly right?: number;
}

const POINTS_PER_PIXEL = 0.75;

/**
 * Layout-only insets for the body paragraphs and tables of every container with a registered
 * inset, keyed by node ID. Nested containers add their insets. Content inside table cells
 * follows its table and is not inset itself.
 */
export function containerInsets(
  part: OoxmlPart,
  index: BlockIndex,
  insetOf: (type: string) => ReactBlockInset | undefined
): Map<string, BlockInset> {
  const out = new Map<string, BlockInset>();
  const add = (node: OoxmlNode, left: number, right: number) => {
    const old = out.get(node.id);
    out.set(node.id, { left: (old?.left ?? 0) + left, right: (old?.right ?? 0) + right });
  };
  const walk = (children: readonly OoxmlNode[], left: number, right: number, depth: number) => {
    for (const child of children) {
      if (child.kind === 'paragraph' || child.kind === 'table') {
        add(child, left, right);
        continue;
      }
      const nested = blockStoryContainerChildren(child);
      if (nested && depth < MAX_STORY_SDT_NESTING) walk(nested, left, right, depth + 1);
    }
  };
  for (const block of index.byId.values()) {
    if (block.partName !== part.name) continue;
    const inset = insetOf(block.payload.type);
    const left = Math.max(0, inset?.left ?? 0) * POINTS_PER_PIXEL;
    const right = Math.max(0, inset?.right ?? 0) * POINTS_PER_PIXEL;
    if (!left && !right) continue;
    const control = findNode(part, block.controlId);
    const content = control ? contentControlContentNodeOf(control) : null;
    if (content) walk(content.children, left, right, 0);
  }
  return out;
}
