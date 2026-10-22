import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { encodeBlockPayload, isBlockId, type BlockPayload } from '../blocks/payload-schema';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { JsonObject } from '../data/json';
import { DOMAIN_BLOCK_TAGS } from '../projection/block-tags';
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

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const mdx = (node: TargetAstNode) => node.type === 'mdxJsxFlowElement';

/** A collision-resistant persistent ID that fits Word's 64-character tag limit. */
export const defaultBlockId = (): string => crypto.randomUUID();

/** The application payload a new registered element describes. Unknown tags have none. */
export function insertedPayload(ast: TargetAstNode, id: string): BlockPayload | null {
  const node = domainNodeFromAst(ast);
  const type = node?.type;
  if (!node || typeof type !== 'string' || !DOMAIN_BLOCK_TAGS.has(type)) return null;
  const { type: _type, ...attributes } = node;
  const payload: BlockPayload = { version: 1, id, type, attributes: attributes as JsonObject };
  return payloadMatches(ast, payload) ? payload : null;
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
  allocate: () => string = defaultBlockId
): { success: true; plan: InsertedBlockPlan } | NativeEditFailure {
  const blocks: InsertedBlockBinding[] = [],
    lists: InsertedListBinding[] = [];
  const visit = (
    block: LogicalBlock,
    ast: TargetAstNode,
    path: readonly number[]
  ): NativeEditFailure | null => {
    if (block.kind === 'paragraph' || block.kind === 'row' || block.kind === 'cell') return null;
    if (block.kind === 'block')
      return fail('New opaque blocks need their own native representation.');
    if (block.kind === 'table') {
      const listed = (item: LogicalBlock): boolean =>
        item.kind === 'paragraph' ? !!item.list : 'children' in item && item.children.some(listed);
      return listed(block) ? fail('Lists inside new tables need nested table list paths.') : null;
    }
    if (!mdx(ast)) {
      if (block.kind === 'container' && ast.type === 'blockquote' && path.length === 1) return null;
      return fail(`New ${ast.type} blocks need their own native representation.`);
    }
    const id = allocate();
    if (!isBlockId(id)) return fail('The block ID allocator returned an invalid ID.');
    const payload = insertedPayload(ast, id);
    const definition = payload && DOMAIN_BLOCK_TAGS.get(payload.type);
    if (!payload || !definition || definition.kind !== block.kind)
      return fail('New application blocks need a registered payload type.');
    blocks.push({ path: [...path], payload, label: definition.tag });
    if (block.kind === 'atomic') return null;
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
  value: unknown
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
      if (!node || !mdx(node) || !payloadMatches(node, binding.payload)) return false;
    }
    return true;
  } catch {
    return false;
  }
}
