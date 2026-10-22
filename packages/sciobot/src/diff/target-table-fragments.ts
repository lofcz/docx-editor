import { resolveTargetLists } from './target-lists';
import type { InsertedListBinding } from './inserted-list-types';
import type { NativeEditFailure } from '../contracts/document';
import type { TargetAstNode } from '../projection/parse-target';
import type { LogicalBlock, LogicalGroup } from './logical-block-types';
import { targetAstKey } from './target-key';
import { compileInlineTarget } from './inline-target';
import { targetParagraphHtmlKey } from './target-inline';

/** Syntax aliases retain all authored metadata and header-cell distinctions. */
export function targetTableHeaderKey(block: LogicalGroup): string {
  const own = { ...block.header };
  delete own.type;
  if (['table', 'tr', 'td'].includes(String(own.name))) delete own.name;
  if (Array.isArray(own.attributes) && !own.attributes.length) delete own.attributes;
  if (Array.isArray(own.align) && own.align.every((value) => value === null)) delete own.align;
  return targetAstKey({ ...own, type: block.kind });
}

/** Only new cell fragments use target ASTs. Retained native content never passes through this codec. */
export function insertedTargetCellFragments(
  cell: LogicalGroup,
  targetAst: TargetAstNode,
  astOf: (block: LogicalBlock) => TargetAstNode,
  source?: Parameters<typeof resolveTargetLists>[0]
):
  | {
      readonly success: true;
      readonly fragments: readonly TargetAstNode[];
      readonly lists: readonly InsertedListBinding[];
    }
  | NativeEditFailure {
  const fail = (error: string): NativeEditFailure => ({
    success: false,
    code: 'unsupported-operation',
    error,
  });
  if (targetTableHeaderKey(cell) !== targetAstKey({ type: 'cell' }))
    return fail('Inserted cell metadata needs explicit native properties.');
  if (!source && cell.children.some((block) => block.list))
    return fail('Inserted cell lists need explicit native list destinations.');
  const resolved = source
    ? resolveTargetLists(source, cell.children, [])
    : { success: true as const, targets: [] };
  if (!resolved.success) return resolved;
  const lists = resolved.targets.flatMap((target, index) =>
    target.kind === 'none' ? [] : [{ path: [index], target }]
  );
  const fragments: TargetAstNode[] = [];
  for (const block of cell.children) {
    if (block.kind === 'paragraph') {
      if (targetParagraphHtmlKey(block) !== 'null')
        return fail('Inserted paragraph metadata needs explicit native property intentions.');
      const compiled = compileInlineTarget(block.ast);
      if (!compiled.success) return compiled;
      fragments.push(astOf(block));
    } else {
      let original: TargetAstNode | undefined = targetAst;
      for (const step of block.path)
        original = Array.isArray(original?.children)
          ? (original.children[step] as TargetAstNode | undefined)
          : undefined;
      if (!original) return fail('An inserted cell fragment has no original target AST.');
      fragments.push(original);
    }
  }
  return { success: true, fragments, lists };
}
