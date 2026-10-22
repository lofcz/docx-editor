import type { TargetAstNode } from '../projection/parse-target';
import { isListTarget, type ListTarget } from './list-types';

export interface InsertedListBinding {
  /** Indices start in a sibling edit's inserted array, then follow AST child arrays. */
  readonly path: readonly number[];
  readonly target: ListTarget;
}

/** Validate after the D01 bounded JSON gate. A path never names a native or preview paragraph ID. */
export function isInsertedListBindings(
  inserted: readonly TargetAstNode[],
  bindings: unknown
): bindings is readonly InsertedListBinding[] {
  if (!Array.isArray(bindings)) return false;
  const seen = new Set<string>(),
    instances = new Map<string, string>();
  for (const binding of bindings) {
    if (
      !binding ||
      typeof binding !== 'object' ||
      Array.isArray(binding) ||
      Object.keys(binding).some((key) => !['path', 'target'].includes(key)) ||
      !Array.isArray(binding.path) ||
      !binding.path.length ||
      !binding.path.every((step: unknown) => Number.isSafeInteger(step) && Number(step) >= 0) ||
      !isListTarget(binding.target)
    )
      return false;
    const key = binding.path.join('/');
    if (seen.has(key)) return false;
    seen.add(key);
    let siblings: readonly TargetAstNode[] = inserted,
      node: TargetAstNode | undefined;
    for (const index of binding.path) {
      node = siblings[index];
      if (!node) return false;
      siblings = Array.isArray(node.children) ? (node.children as TargetAstNode[]) : [];
    }
    if (!node || !['paragraph', 'heading'].includes(node.type)) return false;
    const target = binding.target;
    if ((target.kind === 'ordered' || target.kind === 'bullet') && target.instance.kind === 'new') {
      const levelKey = JSON.stringify([target.instance.key, target.level]),
        definition = JSON.stringify([
          target.kind,
          target.kind === 'ordered' ? (target.instance.start ?? 1) : null,
        ]);
      const previous = instances.get(levelKey);
      if (previous !== undefined && previous !== definition) return false;
      instances.set(levelKey, definition);
    }
  }
  return true;
}
