import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { targetSourceRangeAt, type TargetAstNode } from '../projection/parse-target';
import type { LogicalParagraph } from './logical-block-types';
import type { BoundLogicalSource, BoundSourceBlock } from './bound-source-types';
import type { TargetSourceInput } from './target-source-types';
import { compileInlineTarget } from './inline-target';

/** Remove only a cell separator outside its one native paragraph's mapped extent and model text. */
export function withoutGeneratedCellSeparator(
  input: TargetSourceInput,
  source: BoundLogicalSource,
  binding: BoundSourceBlock | undefined,
  before: LogicalParagraph
): LogicalParagraph {
  if (
    !binding?.table?.cellId ||
    before.list ||
    before.prefixRemoved ||
    binding.contributions.length !== 1
  )
    return before;
  const children = before.ast.children as TargetAstNode[],
    first = children?.[0],
    contribution = binding.contributions[0]!;
  if (
    !first ||
    !(
      first.type === 'break' ||
      (first.type === 'mdxJsxTextElement' &&
        first.name === 'br' &&
        Array.isArray(first.attributes) &&
        !first.attributes.length &&
        Array.isArray(first.children) &&
        !first.children.length)
    )
  )
    return before;
  const position = first.position as
    | { start?: { offset?: number }; end?: { offset?: number } }
    | undefined;
  if (
    !Number.isSafeInteger(position?.start?.offset) ||
    !Number.isSafeInteger(position?.end?.offset)
  )
    return before;
  const range = targetSourceRangeAt(
    source.parsed,
    position!.start!.offset!,
    position!.end!.offset!
  );
  if (!range || range.end > contribution.start) return before;
  const ast = { ...before.ast, children: children.slice(1) },
    compiled = compileInlineTarget(ast),
    part = input.pkg.parts.get(source.index.partName)!;
  return compiled.success &&
    compiled.target.plainText !== null &&
    paragraphTextOf(part, contribution.ownerId) === compiled.target.plainText
    ? { ...before, ast }
    : before;
}
