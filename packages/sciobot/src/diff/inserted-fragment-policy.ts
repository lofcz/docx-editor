import { stylesPartOf, type OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { TargetAstNode } from '../projection/parse-target';
import type { LogicalBlock } from './logical-blocks';
import { compileInlineTarget } from './inline-target';
import type { ListTarget } from './list-types';
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

export const logicalBlockAst = (block: LogicalBlock): TargetAstNode =>
  'ast' in block
    ? block.ast
    : ({
        ...block.header,
        type: block.header.type as string,
        children: block.children.map(logicalBlockAst),
      } as TargetAstNode);

/** The document's own Quote paragraph style ID, which the native projection reads as a quote. */
export function documentQuoteStyle(pkg: OoxmlPackage): string | null {
  const root = stylesPartOf(pkg)?.root;
  if (!root) return null;
  for (const node of root.children) {
    if (node.kind === 'textValue' || node.localName !== 'style') continue;
    const attribute = (name: string) =>
      node.attributes.find((item) => item.localName === name)?.value;
    const id = attribute('styleId');
    if (attribute('type') === 'paragraph' && id && id.toLowerCase() === 'quote') return id;
  }
  return null;
}
const inlineProblem = (ast: TargetAstNode): NativeEditFailure | null => {
  const inline = compileInlineTarget(ast);
  if (!inline.success) return inline;
  // D30 validates each atom and binds payload/resource ownership after structural pairing.
  // Inline code is the application character style (N18); fragment seeds write it.
  return null;
};
/** Inserted blocks with a native fragment representation; everything else refuses before planning. */
export function insertedBlockProblem(
  block: LogicalBlock,
  ast: TargetAstNode,
  list: ListTarget,
  nested = false
): NativeEditFailure | null {
  if (block.kind === 'paragraph') return inlineProblem(ast);
  const registered = ast.type === 'mdxJsxFlowElement';
  if (block.kind === 'block' && ast.type === 'thematicBreak') return null;
  const listed = (item: LogicalBlock): boolean =>
    item.kind === 'paragraph' ? !!item.list : 'children' in item && item.children.some(listed);
  if (list.kind !== 'none' || (!registered && listed(block)))
    return fail('List membership of new rich blocks needs nested list paths.');
  if (nested && ast.type === 'blockquote')
    return fail('Block quotes inside new application blocks need nested style bindings.');
  const children = (ast.children ?? []) as TargetAstNode[];
  if (ast.type === 'blockquote') {
    if (!children.length || children.some((child) => child.type !== 'paragraph'))
      return fail('New block quotes can contain paragraphs only.');
    for (const child of children) {
      const problem = inlineProblem(child);
      if (problem) return problem;
    }
    return null;
  }
  if (block.kind === 'block' && ast.type === 'math') return null;
  if (block.kind === 'atomic' && ast.type === 'mdxJsxFlowElement') return null;
  if (block.kind === 'container' && ast.type === 'mdxJsxFlowElement') {
    // Registered containers: planInsertedBlocks validates payloads; children follow the same rules.
    for (const [index, child] of block.children.entries()) {
      const problem = insertedBlockProblem(
        child,
        ((ast.children ?? []) as TargetAstNode[])[index]!,
        { kind: 'none' },
        true
      );
      if (problem) return problem;
    }
    return null;
  }
  if (ast.type === 'table' && block.kind === 'table') {
    for (const row of children)
      for (const cell of (row.children ?? []) as TargetAstNode[])
        for (const content of (cell.children ?? []) as TargetAstNode[]) {
          if (content.type !== 'paragraph')
            return fail('New table cells can contain paragraphs only.');
          const problem = inlineProblem(content);
          if (problem) return problem;
        }
    return null;
  }
  return fail('Inserted rich blocks still need canonical insertion intentions.');
}
