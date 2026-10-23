import type { TargetAstNode } from '../projection/parse-target';
import { compileInlineTarget } from './inline-target';
import { compileInsertedAtoms, type InsertedInlineAtom } from './inserted-atoms';
import { jsonDataKey } from './target-key';

/** Remove only resolved atom paths. Literal text, marks and authored link boundaries remain. */
export function withoutInlineAtoms(ast: TargetAstNode): TargetAstNode | null {
  const compiled = compileInlineTarget(ast);
  if (!compiled.success) return null;
  const atoms = new Set(
    compiled.target.tokens
      .filter((token) => token.kind === 'atom')
      .map((token) => token.path.join('/'))
  );
  const visit = (node: TargetAstNode, path: number[]): TargetAstNode | null => {
    if (atoms.has(path.join('/'))) return null;
    if (!Array.isArray(node.children)) return node;
    const children = (node.children as TargetAstNode[]).flatMap((child, index) => {
      const next = visit(child, [...path, index]);
      return next ? [next] : [];
    });
    if (path.length && !children.length) return null;
    return { ...node, children };
  };
  return visit(ast, []);
}

/** D01 validation reconstructs the full intentions without invoking native writers. */
export function isRetainedAtomBindings(
  paragraphId: string,
  offset: number,
  before: string,
  ast: TargetAstNode,
  value: unknown
): boolean {
  if (!Array.isArray(value) || !value.length) return false;
  try {
    const controls = (value as InsertedInlineAtom[]).filter((op) => op.kind === 'inline-control');
    const ids = controls.map((op) => op.payload.id);
    if (new Set(ids).size !== ids.length) return false;
    let next = 0;
    const compiled = compileInsertedAtoms(
      {
        location: { kind: 'retained-paragraph', paragraphId },
        offset,
        allocateBlockId: () => ids[next++] ?? '',
      },
      ast
    );
    if (!compiled.success || next !== controls.length) return false;
    const skeleton = withoutInlineAtoms(ast);
    const text = skeleton && compileInlineTarget(skeleton);
    return (
      !!text &&
      text.success &&
      text.target.plainText === before &&
      jsonDataKey(compiled.plan.operations as unknown as import('../data/json').JsonValue) ===
        jsonDataKey(value)
    );
  } catch {
    return false;
  }
}
