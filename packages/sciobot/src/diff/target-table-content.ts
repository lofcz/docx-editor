import type { LogicalBlock } from './logical-block-types';
import type { TargetTableContext } from './target-table';
import { targetParagraphHtmlKey } from './target-inline';
import { targetAstKey } from './target-key';
import { targetTableHeaderKey } from './target-table-fragments';
import { withoutGeneratedCellSeparator } from './target-generated-separators';

/** Preserve native content across transparent HTML syntax and a source-proven generated cell separator. */
export function sameTargetCellContent(
  context: TargetTableContext,
  before: readonly LogicalBlock[],
  after: readonly LogicalBlock[],
  path: readonly number[]
): boolean {
  if (before.length !== after.length) return false;
  return before.every((old, index) => {
    const next = after[index]!;
    if (context.same(old, next)) return true;
    if (old.kind !== next.kind) return false;
    if ('children' in old && 'children' in next) {
      if (
        old.domainType !== next.domainType ||
        targetTableHeaderKey(old) !== targetTableHeaderKey(next)
      )
        return false;
      const groups = (block: typeof old) =>
        targetAstKey({
          type: 'groups',
          children: block.rowGroups.map((group) => ({
            ...group.header,
            type: 'group',
            start: group.start,
            end: group.end,
          })),
        });
      return (
        groups(old) === groups(next) &&
        sameTargetCellContent(context, old.children, next.children, [...path, index])
      );
    }
    if (
      old.kind !== 'paragraph' ||
      next.kind !== 'paragraph' ||
      old.list ||
      next.list ||
      old.prefixRemoved ||
      next.prefixRemoved ||
      targetParagraphHtmlKey(old) !== targetParagraphHtmlKey(next)
    )
      return false;
    const normalized = withoutGeneratedCellSeparator(
      context.input,
      context.source,
      context.bindings.get([...path, index].join('/')),
      old
    );
    return normalized !== old && targetAstKey(normalized.ast) === targetAstKey(next.ast);
  });
}

/**
 * A native cell keeps one paragraph. Emptying a cell therefore retains its first paragraph with
 * no text, which native merges then treat as blank, instead of deleting every block.
 */
export function blankCellTarget(
  before: readonly LogicalBlock[],
  after: readonly LogicalBlock[],
  path: readonly number[]
): readonly LogicalBlock[] {
  if (after.length || !before.length || before[0]!.kind !== 'paragraph') return after;
  return [
    {
      kind: 'paragraph',
      path,
      list: null,
      key: 'paragraph:blank-cell',
      ast: { type: 'paragraph', children: [] },
      synthetic: true,
      inlinePaths: [],
      prefixRemoved: 0,
      html: null,
    },
  ];
}
