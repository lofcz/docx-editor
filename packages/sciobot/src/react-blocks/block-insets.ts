import type { BlockInset, CollapsedContentControl } from '@docx-editor.dev/core/editor';
import {
  blockStoryContainerChildren,
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
  insetOf: (type: string) => ReactBlockInset | undefined,
  collapsed: readonly CollapsedContentControl[] = []
): Map<string, BlockInset> {
  const out = new Map<string, BlockInset>();
  const folds = new Map(collapsed.map((entry) => [entry.controlId, entry]));
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
        if (!nested || depth >= MAX_STORY_SDT_NESTING) continue;
        const fold = folds.get(child.id);
        if (fold) {
          if (fold.placeholderHeightPt === 0) continue;
          const anchor = fold.summaryParagraphId
            ? findNode(part, fold.summaryParagraphId)
            : firstParagraph(nested, depth + 1);
          if (anchor?.kind === 'paragraph') into.push(anchor);
        } else blocks(nested, into, depth + 1);
      }
    }
    return into;
  };
  const firstParagraph = (children: readonly OoxmlNode[], depth: number): OoxmlNode | undefined => {
    for (const child of children) {
      if (child.kind === 'paragraph' || child.kind === 'table') return child;
      const nested = blockStoryContainerChildren(child);
      if (nested && depth < MAX_STORY_SDT_NESTING) {
        const first = firstParagraph(nested, depth + 1);
        if (first) return first;
      }
    }
    return undefined;
  };
  for (const block of index.byId.values()) {
    if (block.partName !== part.name) continue;
    const inset = insetOf(block.payload.type);
    const points = (value: number | undefined) => Math.max(0, value ?? 0) * POINTS_PER_PIXEL;
    const room = {
      left: points(inset?.left),
      right: points(inset?.right),
      top: folds.get(block.controlId)?.summaryParagraphId === null ? 0 : points(inset?.top),
      bottom: folds.get(block.controlId)?.summaryParagraphId === null ? 0 : points(inset?.bottom),
    };
    if (!room.left && !room.right && !room.top && !room.bottom) continue;
    const control = findNode(part, block.controlId);
    const members = control ? blocks([control], [], 0) : [];
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
