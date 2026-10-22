import type { MdxProjection } from '../projection/mdx';
import {
  parseTargetMarkdown,
  targetSourceRangeAt,
  type ParsedTargetMarkdown,
  type TargetAstNode,
} from '../projection/parse-target';
import { targetAstKey } from '../diff/target-key';
interface TableSyntax {
  readonly key: string;
  readonly start: number;
  readonly end: number;
}
function tables(ast: TargetAstNode): TableSyntax[] {
  const found: TableSyntax[] = [];
  const visit = (node: TargetAstNode) => {
    if (node.type === 'table') {
      const position = node.position as
        | { start?: { offset?: number }; end?: { offset?: number } }
        | undefined;
      if (position?.start?.offset !== undefined && position.end?.offset !== undefined)
        found.push({
          key: targetAstKey(node),
          start: position.start.offset,
          end: position.end.offset,
        });
    }
    if (Array.isArray(node.children))
      for (const child of node.children as TargetAstNode[]) visit(child);
  };
  visit(ast);
  return found;
}
/**
 * An unchanged GFM table is read syntax, not an instruction to replace native paragraph boundaries
 * with its generated <br> separators. Resolve only unique complete table matches to the native
 * auto syntax. Changed GFM cells still pass through the normal planner and its existing limits.
 */
export function preserveUnchangedTableSyntax(
  target: ParsedTargetMarkdown,
  gfm: MdxProjection,
  auto: MdxProjection
): string {
  const parsed = parseTargetMarkdown(gfm.text);
  if (!parsed.success) return target.markdown;
  const current = tables(parsed.target.ast),
    requested = tables(target.ast),
    counts = (entries: readonly TableSyntax[]) => {
      const map = new Map<string, number>();
      for (const entry of entries) map.set(entry.key, (map.get(entry.key) ?? 0) + 1);
      return map;
    },
    oldCounts = counts(current),
    newCounts = counts(requested),
    replacements: { start: number; end: number; text: string }[] = [];
  for (const table of requested) {
    if (oldCounts.get(table.key) !== 1 || newCounts.get(table.key) !== 1) continue;
    const old = current.find((entry) => entry.key === table.key)!;
    const sourceRange = targetSourceRangeAt(parsed.target, old.start, old.end);
    const range = gfm.sourceMap.tables.find(
      (range) =>
        range.sourceScope === 'body' &&
        range.markdownStart === sourceRange?.start &&
        range.markdownEnd === sourceRange?.end
    );
    const native =
      range &&
      auto.sourceMap.tables.find(
        (candidate) => candidate.tableId === range.tableId && candidate.sourceScope === 'body'
      );
    if (!native) continue;
    const text = auto.text.slice(native.markdownStart, native.markdownEnd);
    if (text.trimStart().startsWith('<table'))
      replacements.push({ start: table.start, end: table.end, text });
  }
  let markdown = target.markdown;
  for (const replacement of replacements.sort((a, b) => b.start - a.start))
    markdown =
      markdown.slice(0, replacement.start) + replacement.text + markdown.slice(replacement.end);
  return markdown;
}
