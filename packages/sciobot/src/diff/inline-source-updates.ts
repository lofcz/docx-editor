import { activeMdxRegistry } from '../registry/active';
import { updatableInlineComponent } from '../registry/mdx-read';
import { buildBlockIndex } from '../blocks/block-index';
import { targetSourceRangeAt, type TargetAstNode } from '../projection/parse-target';
import { planAtomicBlock } from './atomic-blocks';
import { domainNodeFromAst } from './domain-ast';
import { alignRetainedAtoms } from './retained-atom-alignment';
import { retainedAtomKey } from './retained-atom-gaps';
import { jsonDataKey } from './target-key';
import type { CompiledInlineTarget } from './inline-target';
import type { TargetParagraphSource } from './target-inline';
import type { CanonicalEdit, EditPlanResult } from './plan';

/** Source property changes retain their original native control, even beside atom births/removals. */
export function planInlineSourceUpdates(
  context: TargetParagraphSource,
  before: CompiledInlineTarget,
  after: CompiledInlineTarget
): { context: TargetParagraphSource; edits: readonly CanonicalEdit[] } | EditPlanResult | null {
  if (context.nativeRange) return null;
  const old = before.tokens.filter((token) => token.kind === 'atom'),
    next = after.tokens.filter((token) => token.kind === 'atom');
  const registry = activeMdxRegistry();
  const updatable = (ast: TargetAstNode) => updatableInlineComponent(ast, registry);
  if (!old.some((token) => updatable(token.ast))) return null;
  const keys = new Map<string, number>();
  for (const token of next)
    keys.set(retainedAtomKey(token.ast), (keys.get(retainedAtomKey(token.ast)) ?? 0) + 1);
  let changed = false;
  for (const token of old)
    if (updatable(token.ast)) {
      const key = retainedAtomKey(token.ast),
        count = keys.get(key) ?? 0;
      if (count) keys.set(key, count - 1);
      else changed = true;
    }
  if (!changed) return null;
  const alignment =
    context.retainedAtomAlignment ?? alignRetainedAtoms(before, after, true, false, true);
  if (!alignment) return null;
  const edits: CanonicalEdit[] = [],
    replacements = new Map<string, TargetAstNode>();
  const blocks = buildBlockIndex(context.input.pkg);
  const fail = (error: string): EditPlanResult => ({
    success: false,
    code: 'unsupported-operation',
    error,
  });
  for (let i = 0; i < old.length; i++) {
    const atom = old[i]!,
      target = next[alignment[i]!];
    const component = updatable(atom.ast);
    if (
      !component ||
      !target ||
      updatable(target.ast) !== component ||
      retainedAtomKey(atom.ast) === retainedAtomKey(target.ast)
    )
      continue;
    const position = atom.ast.position as
      | { start?: { offset?: number }; end?: { offset?: number } }
      | undefined;
    if (position?.start?.offset === undefined || position.end?.offset === undefined)
      return fail('The changed source has no bound native range.');
    const range = targetSourceRangeAt(
      context.source.parsed,
      position.start.offset,
      position.end.offset
    );
    const bindings =
      range &&
      context.binding.inline.filter(
        (entry) => entry.start === range.start && entry.end === range.end
      );
    if (!bindings || bindings.length !== 1)
      return fail('The changed source has no unique native control.');
    const block = blocks.byId.get(bindings[0]!.blockId),
      a = domainNodeFromAst(atom.ast, registry),
      b = domainNodeFromAst(target.ast, registry);
    if (
      !block ||
      block.payload.type !== component.type ||
      a?.type !== component.type ||
      b?.type !== component.type
    )
      return fail('The changed source needs its registered native codec.');
    const attributes = { ...block.payload.attributes };
    for (const name of new Set([...Object.keys(a), ...Object.keys(b)])) {
      if (
        name === 'type' ||
        (jsonDataKey(a[name] ?? null) === jsonDataKey(b[name] ?? null) &&
          Object.hasOwn(a, name) === Object.hasOwn(b, name))
      )
        continue;
      if (Object.hasOwn(b, name))
        Object.defineProperty(attributes, name, {
          value: b[name],
          enumerable: true,
          configurable: true,
          writable: true,
        });
      else delete attributes[name];
    }
    const planned = planAtomicBlock(
      { revision: context.source.index.revision, pkg: context.input.pkg, blockId: block.id },
      { ...block.payload, attributes },
      registry
    );
    if (!planned.success) return planned;
    edits.push(...planned.plan.operations);
    replacements.set(atom.path.join('/'), { ...target.ast, position: atom.ast.position });
  }
  if (!replacements.size) return null;
  const visit = (node: TargetAstNode, path: number[]): TargetAstNode =>
    replacements.get(path.join('/')) ??
    (node.children
      ? {
          ...node,
          children: (node.children as TargetAstNode[]).map((child, i) =>
            visit(child, [...path, i])
          ),
        }
      : node);
  return {
    edits,
    context: {
      ...context,
      before: { ...context.before, ast: visit(before.ast, []) },
      retainedAtomAlignment: alignment,
    },
  };
}
