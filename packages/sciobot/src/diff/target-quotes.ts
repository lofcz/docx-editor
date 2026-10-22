import type { BoundLogicalSource, BoundSourceBlock } from './bound-source-types';
import type { LogicalBlock, LogicalBlockDocument } from './logical-block-types';
import type { TargetAstNode } from '../projection/parse-target';

/** A Word quote consists of sibling paragraphs, rather than a native container node. */
export function quoteParagraphStream(document: LogicalBlockDocument) {
  const blocks: LogicalBlock[] = [],
    quoted = new Set<LogicalBlock>(),
    origins = new Map<string, readonly number[]>();
  const origin = (block: LogicalBlock, physical: readonly number[], logical: readonly number[]) => {
    origins.set(physical.join('/'), logical);
    if ('children' in block)
      block.children.forEach((child, index) =>
        origin(child, [...physical, index], [...logical, index])
      );
  };
  for (const [index, block] of document.blocks.entries()) {
    if (
      block.kind === 'container' &&
      block.header.type === 'blockquote' &&
      block.children.length &&
      block.children.every(
        (child) => child.kind === 'paragraph' && child.ast.type === 'paragraph' && !child.list
      )
    ) {
      for (const [childIndex, child] of block.children.entries()) {
        origin(child, [blocks.length], [index, childIndex]);
        quoted.add(child);
        blocks.push(child);
      }
    } else {
      origin(block, [blocks.length], [index]);
      blocks.push(block);
    }
  }
  return { document: { ...document, blocks }, quoted, origins };
}

/** Keep each paragraph's original source binding; never fabricate a quote owner. */
export function bindQuoteParagraphStream(source: BoundLogicalSource) {
  const stream = quoteParagraphStream(source.logical),
    original = new Map(source.bindings.map((binding) => [binding.logicalPath.join('/'), binding])),
    bindings: BoundSourceBlock[] = [];
  for (const [physical, logical] of stream.origins) {
    const binding = original.get(logical.join('/'));
    if (binding) bindings.push({ ...binding, logicalPath: physical.split('/').map(Number) });
  }
  const physical = new Map(
    [...stream.origins].map(([path, logical]) => [logical.join('/'), path.split('/').map(Number)])
  );
  const windows = source.windows.map((window) => ({
    ...window,
    members: window.members.map((member) => ({
      ...member,
      logicalPaths: member.logicalPaths.flatMap((path) => {
        const mapped = physical.get(path.join('/'));
        return mapped ? [mapped] : [];
      }),
    })),
  }));
  return { ...stream, source: { ...source, logical: stream.document, bindings, windows } };
}

/** Preserve one fragment for adjacent new paragraphs from the same logical quote. */
export function quoteSiblingInsertion(
  nodes: readonly TargetAstNode[],
  groups: readonly (string | null)[]
) {
  const inserted: TargetAstNode[] = [],
    positions: number[] = [];
  let previous: string | null = null;
  for (const [index, node] of nodes.entries()) {
    const group = groups[index] ?? null;
    if (group !== null && group === previous) {
      const last = inserted.at(-1)!;
      inserted[inserted.length - 1] = {
        ...last,
        children: [...(last.children as TargetAstNode[]), node],
      };
    } else
      inserted.push(
        group === null ? node : ({ type: 'blockquote', children: [node] } as TargetAstNode)
      );
    positions.push(inserted.length - 1);
    previous = group;
  }
  return { inserted, path: (path: readonly number[]) => [positions[path[0]!]!, ...path.slice(1)] };
}
