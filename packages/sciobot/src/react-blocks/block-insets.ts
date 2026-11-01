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
  /** Space above the container's first paragraph, at least. */
  readonly top?: number;
  /** Space below the container's last paragraph, at least. */
  readonly bottom?: number;
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
  const add = (node: OoxmlNode, room: Required<ReactBlockInset>) => {
    const old = out.get(node.id);
    const top = (old?.top ?? 0) + room.top,
      bottom = (old?.bottom ?? 0) + room.bottom;
    out.set(node.id, {
      left: (old?.left ?? 0) + room.left,
      right: (old?.right ?? 0) + room.right,
      ...(top ? { top } : {}),
      ...(bottom ? { bottom } : {}),
    });
  };
  const blocks = (children: readonly OoxmlNode[], into: OoxmlNode[], depth: number) => {
    for (const child of children) {
      if (child.kind === 'paragraph' || child.kind === 'table') into.push(child);
      else {
        const nested = blockStoryContainerChildren(child);
        if (nested && depth < MAX_STORY_SDT_NESTING) blocks(nested, into, depth + 1);
      }
    }
    return into;
  };
  for (const block of index.byId.values()) {
    if (block.partName !== part.name) continue;
    const inset = insetOf(block.payload.type);
    const points = (value: number | undefined) => Math.max(0, value ?? 0) * POINTS_PER_PIXEL;
    const room = {
      left: points(inset?.left),
      right: points(inset?.right),
      top: points(inset?.top),
      bottom: points(inset?.bottom),
    };
    if (!room.left && !room.right && !room.top && !room.bottom) continue;
    const control = findNode(part, block.controlId);
    const content = control ? contentControlContentNodeOf(control) : null;
    const members = content ? blocks(content.children, [], 0) : [];
    // Top and bottom room belong to the first and last paragraph; a table edge takes none.
    members.forEach((member, at) =>
      add(member, {
        left: room.left,
        right: room.right,
        top: at === 0 && member.kind === 'paragraph' ? room.top : 0,
        bottom: at === members.length - 1 && member.kind === 'paragraph' ? room.bottom : 0,
      })
    );
  }
  return out;
}
