import { MDX_BLOCK } from '../projection/mdx';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { encodeBlockPayload, isBlockId, type BlockPayload } from '../blocks/payload-schema';
import type { MdxComponentRegistry } from '../registry/contract';
import { sciobotMdxRegistry } from '../registry/sciobot';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { JsonObject } from '../data/json';
import type { TargetAstNode } from '../projection/parse-target';
import { domainNodeFromAst } from './domain-ast';
import { payloadMatches } from './inserted-cell-fragments';
import { insertedPayloadProblem } from './inserted-payloads';
import type { InsertedListBinding } from './inserted-list-types';
import type { LogicalBlock, LogicalGroup } from './logical-block-types';
import { resolveTargetLists } from './target-lists';

/** One new registered block inside a sibling insertion; its ID is new and verified unique. */
export interface InsertedBlockBinding {
  /** Inserted index, then AST child indices. */
  readonly path: readonly number[];
  readonly payload: BlockPayload;
  /** Readable control alias and payload record label. */
  readonly label: string;
}

export type NewBlockPayload = (
  ast: TargetAstNode,
  id: string
) => { readonly payload: BlockPayload; readonly label: string } | null;

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const mdx = (node: TargetAstNode) => node.type === 'mdxJsxFlowElement';
/** A Markdown node that a registered block component reads, such as fenced code. */
const markdownBlock = (registry: MdxComponentRegistry, node: TargetAstNode) =>
  registry.markdown.get(node.type)?.kind === 'block';

/** A collision-resistant persistent ID that fits Word's 64-character tag limit. */
export const defaultBlockId = (): string => crypto.randomUUID();

/** The application payload a new registered element describes. Unknown tags have none. */
export function insertedPayload(
  ast: TargetAstNode,
  id: string,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): BlockPayload | null {
  const markdown = registry.markdown.get(ast.type)?.markdown;
  if (markdown) return markdown.fromMarkdown(ast, id);
  const node = domainNodeFromAst(ast, registry);
  const type = node?.type;
  if (!node || typeof type !== 'string' || !registry.blockTags.has(type)) return null;
  const { type: _type, ...attributes } = node;
  const payload: BlockPayload = { version: 1, id, type, attributes: attributes as JsonObject };
  return payloadMatches(ast, payload, registry) ? payload : null;
}

export interface InsertedBlockPlan {
  readonly blocks: readonly InsertedBlockBinding[];
  /** List membership of paragraphs nested inside new registered containers. */
  readonly lists: readonly InsertedListBinding[];
}

/**
 * Bind new registered containers and atomic blocks of one sibling insertion to new payloads, and
 * resolve list groups nested inside them. Pure: native controls and records exist only after an
 * executor applies the plan.
 */
export function planInsertedBlocks(
  source: {
    readonly revision: RevisionToken;
    readonly pkg: OoxmlPackage;
    readonly partName: string;
  },
  inserted: readonly { readonly block: LogicalBlock; readonly ast: TargetAstNode }[],
  allocate: () => string = defaultBlockId,
  authorPayload?: NewBlockPayload,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): { success: true; plan: InsertedBlockPlan } | NativeEditFailure {
  const blocks: InsertedBlockBinding[] = [],
    lists: InsertedListBinding[] = [];
  const visit = (
    block: LogicalBlock,
    ast: TargetAstNode,
    path: readonly number[]
  ): NativeEditFailure | null => {
    if (block.kind === 'paragraph' || block.kind === 'row' || block.kind === 'cell') return null;
    // A block written as a Markdown node, such as fenced code or display math.
    const markdown = block.kind === 'block' && markdownBlock(registry, ast);
    const thematicBreak = block.kind === 'block' && ast.type === 'thematicBreak';
    if (block.kind === 'block' && !markdown && !thematicBreak)
      return fail('New opaque blocks need their own native representation.');
    if (thematicBreak) return null;
    if (block.kind === 'table') {
      const listed = (item: LogicalBlock): boolean =>
        item.kind === 'paragraph' ? !!item.list : 'children' in item && item.children.some(listed);
      return listed(block) ? fail('Lists inside new tables need nested table list paths.') : null;
    }
    if (!mdx(ast) && !markdown) {
      if (block.kind === 'container' && ast.type === 'blockquote' && path.length === 1) return null;
      return fail(`New ${ast.type} blocks need their own native representation.`);
    }
    const id = allocate();
    if (!isBlockId(id)) return fail('The block ID allocator returned an invalid ID.');
    // Markdown components have their own codec; React admission receives MDX syntax only.
    const authored = markdown ? undefined : authorPayload?.(ast, id);
    const payload = authored?.payload ?? insertedPayload(ast, id, registry);
    const definition = payload && registry.blockTags.get(payload.type);
    const opaque = payload?.type === MDX_BLOCK && block.kind === 'atomic';
    if (
      !payload ||
      payload.id !== id ||
      (!opaque && (!definition || definition.kind !== (markdown ? 'atomic' : block.kind))) ||
      !payloadMatches(ast, payload, registry)
    )
      return fail('New application blocks need a registered payload type.');
    const label = authored?.label ?? definition?.tag;
    if (typeof label !== 'string' || label.length > 256)
      return fail('A new application block needs a bounded readable label.');
    blocks.push({ path: [...path], payload, label });
    if (block.kind === 'atomic' || markdown) return null;
    const group = block as LogicalGroup;
    const children = (ast.children ?? []) as TargetAstNode[];
    if (children.length !== group.children.length)
      return fail('A new block does not map its logical children.');
    const resolved = resolveTargetLists(source, group.children, []);
    if (!resolved.success) return resolved;
    for (const [index, child] of group.children.entries()) {
      const target = resolved.targets[index]!;
      if (child.kind === 'paragraph' && target.kind !== 'none')
        lists.push({ path: [...path, index], target });
      const problem = visit(child, children[index]!, [...path, index]);
      if (problem) return problem;
    }
    return null;
  };
  for (const [index, entry] of inserted.entries()) {
    const problem = visit(entry.block, entry.ast, [index]);
    if (problem) return problem;
  }
  const storage = insertedPayloadProblem(
    source.pkg,
    blocks.map((binding) => ({
      path: binding.path,
      payload: binding.payload,
      label: binding.label,
    }))
  );
  if (storage) return storage;
  return { success: true, plan: { blocks, lists } };
}

/** Called after the D01 bounded JSON gate. Each binding names one registered element by path. */
export function isInsertedBlockBindings(
  inserted: readonly TargetAstNode[],
  value: unknown,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): boolean {
  if (!Array.isArray(value)) return false;
  const paths = new Set<string>(),
    ids = new Set<string>();
  try {
    for (const raw of value) {
      if (
        !record(raw) ||
        Object.keys(raw).some((key) => !['path', 'payload', 'label'].includes(key)) ||
        !Array.isArray(raw.path) ||
        !raw.path.length ||
        !raw.path.every((step) => Number.isSafeInteger(step) && Number(step) >= 0) ||
        typeof raw.label !== 'string' ||
        raw.label.length > 256 ||
        !record(raw.payload)
      )
        return false;
      const binding = raw as unknown as InsertedBlockBinding;
      encodeBlockPayload(binding.payload);
      const key = binding.path.join('/');
      if (paths.has(key) || ids.has(binding.payload.id)) return false;
      paths.add(key);
      ids.add(binding.payload.id);
      let siblings: readonly TargetAstNode[] = inserted,
        node: TargetAstNode | undefined;
      for (const step of binding.path) {
        node = siblings[step];
        if (!node) return false;
        siblings = Array.isArray(node.children) ? (node.children as TargetAstNode[]) : [];
      }
      if (
        !node ||
        (!mdx(node) && !markdownBlock(registry, node)) ||
        !payloadMatches(node, binding.payload, registry)
      )
        return false;
    }
    // A complete atomic AST owns its descendants. Extra child bindings would create orphan payloads.
    const byPath = new Map(
      (value as InsertedBlockBinding[]).map((binding) => [binding.path.join('/'), binding])
    );
    let matched = 0;
    const visit = (node: TargetAstNode, path: readonly number[]): void => {
      const binding = byPath.get(path.join('/'));
      if (binding) {
        matched++;
        if (
          binding.payload.type === MDX_BLOCK ||
          registry.blockTags.get(binding.payload.type)?.kind === 'atomic'
        )
          return;
      }
      if (Array.isArray(node.children))
        (node.children as TargetAstNode[]).forEach((child, index) =>
          visit(child, [...path, index])
        );
    };
    inserted.forEach((node, index) => visit(node, [index]));
    return matched === byPath.size;
  } catch {
    return false;
  }
}
