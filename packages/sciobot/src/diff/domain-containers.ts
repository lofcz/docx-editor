import type { MdxComponentRegistry } from '../registry/contract';
import { activeMdxRegistry } from '../registry/active';
import {
  contentControlContentNodeOf,
  findNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import type { TargetAstNode } from '../projection/parse-target';
import { domainNodeFromAst } from './domain-ast';
import { domainDiffStrategy, sameContainerIdentity } from './domain-strategies';
import { createEditPlan, type CanonicalEdit, type CanonicalEditPlan } from './plan';
import { isTargetFragment } from './target-fragment';
import { planSiblingBlocks, siblingAstKey, SIBLING_BLOCK_LIMITS } from './sibling-blocks';

export { domainNodeFromAst } from './domain-ast';
export interface DomainSiblingWindow {
  readonly parentId: string;
  readonly startIndex: number;
  readonly blocks: readonly MappedDomainBlock[];
}
export interface MappedDomainBlock {
  readonly id: string;
  readonly ast: TargetAstNode;
  /** Complete projected children in a contiguous native content window. Labels can remain outside it. */
  readonly children?: DomainSiblingWindow;
}
export interface DomainContainerSource extends DomainSiblingWindow {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  readonly partName: string;
}
export interface DomainContainerPair {
  readonly id: string;
  readonly blockId: string | null;
  readonly parentId: string;
  readonly type: string;
  readonly kind: 'container' | 'content' | 'atomic';
  readonly beforePath: readonly number[];
  readonly afterPath: readonly number[];
  readonly changed: boolean;
}
export type DomainContainerResult =
  | {
      readonly success: true;
      readonly plan: CanonicalEditPlan;
      readonly pairs: readonly DomainContainerPair[];
    }
  | NativeEditFailure;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/** Pair retained domain wrappers recursively. Content/property/atomic pairs stay explicit work for D13-D15. */
export function planDomainContainers(
  source: DomainContainerSource,
  target: readonly TargetAstNode[],
  registry: MdxComponentRegistry = activeMdxRegistry()
): DomainContainerResult {
  if (!source || typeof source !== 'object') return fail('Missing domain container source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'pkg', 'partName', 'parentId', 'startIndex', 'blocks'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Domain source fields must be plain values.');
  const data = {
    revision: source.revision,
    partName: source.partName,
    parentId: source.parentId,
    startIndex: source.startIndex,
    blocks: source.blocks,
    target,
  };
  const problem = jsonDataProblem(data, SIBLING_BLOCK_LIMITS);
  if (problem) return fail(problem);
  const raw = JSON.stringify(data);
  if (
    raw.length > SIBLING_BLOCK_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > SIBLING_BLOCK_LIMITS.maxBytes
  )
    return fail('Domain input exceeds its byte limit.');
  const part = source.pkg?.parts.get(source.partName);
  if (!part) return fail('The native domain part does not exist.', 'invalid-document');
  const index = buildBlockIndex(source.pkg),
    byControl = new Map(
      [...index.byId.values()]
        .filter((block) => block.partName === part.name)
        .map((block) => [block.controlId, block])
    );
  const pairs: DomainContainerPair[] = [],
    operations: CanonicalEdit[] = [];
  let visits = 0;
  const key = (ast: TargetAstNode) => siblingAstKey(ast);
  const walk = (
    window: DomainSiblingWindow,
    after: readonly TargetAstNode[],
    oldPath: readonly number[],
    newPath: readonly number[],
    depth: number
  ): NativeEditFailure | null => {
    if (++visits > 100_000 || depth > 64) return fail('Domain recursion exceeds its limit.');
    if (!Array.isArray(window.blocks)) return fail('Invalid domain child window.');
    const mappedBlocks: readonly MappedDomainBlock[] = window.blocks;
    if (
      mappedBlocks.some(
        (mapped) =>
          !mapped ||
          typeof mapped !== 'object' ||
          Array.isArray(mapped) ||
          typeof mapped.id !== 'string' ||
          !isTargetFragment(mapped.ast) ||
          Object.keys(mapped).some((key) => !['id', 'ast', 'children'].includes(key)) ||
          (mapped.children !== undefined &&
            (!mapped.children ||
              typeof mapped.children !== 'object' ||
              Array.isArray(mapped.children) ||
              !Array.isArray(mapped.children.blocks)))
      )
    )
      return fail('Invalid mapped domain block or child window.');
    const owners = new Map(window.blocks.map((mapped) => [mapped.ast, byControl.get(mapped.id)]));
    for (const mapped of window.blocks) {
      const owner = byControl.get(mapped.id),
        semantic = domainNodeFromAst(mapped.ast, registry);
      const definition = owner && registry.blockTags.get(owner.payload.type);
      if (semantic && !owner)
        return fail('The mapped domain wrapper has no native payload.', 'invalid-document');
      if (definition && mapped.ast.name !== definition.tag)
        return fail('The mapped domain tag does not match its native payload.', 'invalid-document');
      if (
        owner &&
        semantic &&
        definition &&
        semantic.type !== owner.payload.type &&
        // A container that shares its tag with an atom reads as the atom without text.
        registry.components.get(owner.payload.type)?.parse?.tag !== mapped.ast.name
      )
        return fail(
          'The mapped domain type does not match its native payload.',
          'invalid-document'
        );
    }
    const paired = planSiblingBlocks({ ...window, revision: source.revision, part }, after, {
      key,
      pairable(before, next) {
        const owner = owners.get(before),
          nativeStrategy = owner && domainDiffStrategy({ type: owner.payload.type }, registry);
        if (owner && (!nativeStrategy || nativeStrategy.kind === 'atomic'))
          return before.type === next.type;
        const decoded = domainNodeFromAst(before, registry),
          a = decoded && owner ? { ...decoded, type: owner.payload.type } : decoded,
          b = domainNodeFromAst(next, registry);
        if (!a && !b)
          return (
            before.type === next.type &&
            ['paragraph', 'heading', 'code', 'blockquote', 'list', 'listItem', 'table'].includes(
              before.type
            )
          );
        if (!a || !b || a.type !== b.type) return false;
        const strategy = domainDiffStrategy(a, registry);
        return strategy?.kind === 'atomic' || sameContainerIdentity(a, b, registry);
      },
    });
    if (!paired.success) return paired;
    operations.push(...paired.plan.operations);
    for (const pair of paired.pairs) {
      const mapped: MappedDomainBlock = mappedBlocks[pair.beforeIndex]!,
        ast = mapped.ast,
        next = after[pair.afterIndex]!;
      const block = byControl.get(mapped.id),
        before = domainNodeFromAst(ast, registry),
        strategy = block
          ? domainDiffStrategy({ type: block.payload.type }, registry)
          : before && domainDiffStrategy(before, registry);
      const kind =
        strategy?.kind === 'container'
          ? 'container'
          : block || strategy?.kind === 'atomic'
            ? 'atomic'
            : 'content';
      const beforePath = [...oldPath, pair.beforeIndex],
        afterPath = [...newPath, pair.afterIndex];
      pairs.push(
        Object.freeze({
          id: mapped.id,
          blockId: block?.id ?? null,
          parentId: window.parentId,
          type: block?.payload.type ?? ast.type,
          kind,
          beforePath: Object.freeze(beforePath),
          afterPath: Object.freeze(afterPath),
          changed: pair.changed,
        })
      );
      if (kind === 'container') {
        if (!pair.changed && !mapped.children) continue;
        const content = contentControlContentNodeOf(findNode(part, mapped.id)!);
        const children = mapped.children,
          oldChildren = Array.isArray(ast.children) ? (ast.children as TargetAstNode[]) : [],
          newChildren = Array.isArray(next.children) ? (next.children as TargetAstNode[]) : [];
        if (
          !children &&
          oldChildren.length === newChildren.length &&
          oldChildren.every((child, at) => key(child) === key(newChildren[at]!))
        )
          continue;
        if (
          !content ||
          !children ||
          children.parentId !== content.id ||
          children.blocks.length !== oldChildren.length ||
          children.blocks.some((child, at) => key(child.ast) !== key(oldChildren[at]!))
        )
          return fail(
            'A retained domain container needs its exact native child window.',
            'ambiguous-target'
          );
        const error = walk(children, newChildren, beforePath, afterPath, depth + 1);
        if (error) return error;
      } else if (mapped.children)
        return fail('Only retained domain containers may map recursive children.');
    }
    return null;
  };
  const copy = JSON.parse(raw) as typeof data;
  const error = walk(copy, copy.target, [], [], 0);
  if (error) return error;
  const planned = createEditPlan(source.revision, operations);
  return planned.success
    ? { success: true, plan: planned.plan, pairs: Object.freeze(pairs) }
    : planned;
}
