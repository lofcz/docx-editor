import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { MdxComponentRegistry } from '../registry/contract';
import { sciobotMdxRegistry } from '../registry/sciobot';
import { MDX_BLOCK, MDX_INLINE, serializeOpaqueMdx } from '../projection/mdx';
import type { TargetAstNode } from '../projection/parse-target';
import { domainNodeFromAst } from './domain-ast';
import { isTargetFragment } from './target-fragment';
import { jsonDataKey, targetAstKey } from './target-key';

/** Indices address a new cell stream, its block, and nested AST children. */
export interface InsertedCellPayloadBinding {
  readonly path: readonly number[];
  readonly payload: BlockPayload;
  readonly label: string;
}

const blocks = new Set([
  'paragraph',
  'heading',
  'blockquote',
  'list',
  'code',
  'table',
  'thematicBreak',
  'math',
  'html',
  'mdxJsxFlowElement',
  'mdxFlowExpression',
]);
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const only = (value: object, keys: readonly string[]) =>
  Object.keys(value).every((key) => keys.includes(key));
const mdx = (node: TargetAstNode) =>
  node.type === 'mdxJsxFlowElement' || node.type === 'mdxJsxTextElement';
const opaque = (payload: BlockPayload) => payload.type === MDX_BLOCK || payload.type === MDX_INLINE;

/** Whether a target node still describes `payload` exactly, under the registry's MDX rules. */
export function payloadMatches(
  node: TargetAstNode,
  payload: BlockPayload,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): boolean {
  const markdown = registry.components.get(payload.type)?.markdown;
  if (markdown) return markdown.matches(node, payload);
  if (opaque(payload)) {
    serializeOpaqueMdx(payload);
    return targetAstKey(payload.mdast) === targetAstKey(node);
  }
  const definition = registry.blockTags.get(payload.type);
  if (!definition) return false;
  const attributes = definition
    .attributes(payload)
    .filter((entry) => entry.value !== null)
    .map((entry) => ({
      type: 'mdxJsxAttribute',
      name: entry.name,
      value:
        entry.value === true
          ? null
          : typeof entry.value === 'string'
            ? entry.value
            : JSON.stringify(entry.value),
    }));
  const expected = domainNodeFromAst(
    { type: node.type, name: definition.tag, attributes, children: [] },
    registry
  );
  const actual = domainNodeFromAst(node, registry);
  // The expected element is built without children, so a tag that a container shares with
  // an atom reads as the atom; the payload's own type is the one it must match.
  return (
    actual !== null &&
    expected !== null &&
    jsonDataKey(actual) === jsonDataKey({ ...expected, type: payload.type })
  );
}

/** Call only after the bounded plain-JSON gate. Atomic payloads own their complete AST. */
export function isInsertedCellFragments(
  cells: unknown,
  payloads: unknown,
  inlineAtoms = false,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): boolean {
  if (
    !Array.isArray(cells) ||
    !cells.every(
      (cell) =>
        Array.isArray(cell) &&
        cell.every(
          (node: unknown) => isTargetFragment(node) && blocks.has((node as TargetAstNode).type)
        )
    ) ||
    !Array.isArray(payloads)
  )
    return false;
  const bindings = new Map<string, InsertedCellPayloadBinding>(),
    ids = new Set<string>();
  try {
    for (const raw of payloads) {
      if (
        !record(raw) ||
        !only(raw, ['path', 'payload', 'label']) ||
        !Array.isArray(raw.path) ||
        raw.path.length < 2 ||
        !raw.path.every((step) => Number.isSafeInteger(step) && step >= 0) ||
        typeof raw.label !== 'string' ||
        !record(raw.payload)
      )
        return false;
      const binding = raw as unknown as InsertedCellPayloadBinding;
      encodeBlockPayload(binding.payload);
      const key = binding.path.join('/');
      if (bindings.has(key) || ids.has(binding.payload.id)) return false;
      bindings.set(key, binding);
      ids.add(binding.payload.id);
    }
    let matched = 0;
    const visit = (node: TargetAstNode, path: number[]): boolean => {
      const binding = bindings.get(path.join('/'));
      if (inlineAtoms && !binding && (node.type === 'paragraph' || node.type === 'heading'))
        return true;
      if (mdx(node)) {
        if (!binding || !payloadMatches(node, binding.payload, registry)) return false;
        matched++;
        // Atomic payloads own their complete AST. Nested JSX is not a second native control.
        if (
          opaque(binding.payload) ||
          registry.blockTags.get(binding.payload.type)?.kind === 'atomic'
        )
          return true;
      } else if (binding) {
        if (node.type !== 'code' || !payloadMatches(node, binding.payload, registry)) return false;
        matched++;
        return true;
      }
      return (
        !Array.isArray(node.children) ||
        (node.children as TargetAstNode[]).every((child, index) => visit(child, [...path, index]))
      );
    };
    return (
      (cells as TargetAstNode[][]).every((cell, c) =>
        cell.every((node, b) => visit(node, [c, b]))
      ) && matched === bindings.size
    );
  } catch {
    return false;
  }
}
