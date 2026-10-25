import { BLOCK_MATH, INLINE_MATH, decodeMathPayload } from '../blocks/math-codec';
import { matchesNativeMathAst } from '../projection/native-math';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import { DOMAIN_BLOCK_TAGS } from '../projection/block-tags';
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

export function payloadMatches(node: TargetAstNode, payload: BlockPayload): boolean {
  if (payload.type === 'code_block') {
    return (
      node.type === 'code' &&
      typeof node.value === 'string' &&
      node.value === payload.attributes.value &&
      (node.lang ?? '') === payload.attributes.language
    );
  }
  if (payload.type === BLOCK_MATH || payload.type === INLINE_MATH) {
    const math = decodeMathPayload(encodeBlockPayload(payload));
    if (math.status !== 'supported') return false;
    if (node.type === 'math' || node.type === 'inlineMath')
      return matchesNativeMathAst(node, math.kind, math.source);
    if (
      !mdx(node) ||
      node.name !== (math.kind === 'inline' ? 'scio-inline-equation' : 'scio-equation') ||
      !Array.isArray(node.children) ||
      node.children.length ||
      !Array.isArray(node.attributes) ||
      node.attributes.length !== 1
    )
      return false;
    const attr = node.attributes[0];
    return (
      Object.keys(node).every((key) =>
        ['type', 'name', 'attributes', 'children', 'position'].includes(key)
      ) &&
      !!attr &&
      typeof attr === 'object' &&
      !Array.isArray(attr) &&
      attr.type === 'mdxJsxAttribute' &&
      attr.name === 'texExpression' &&
      attr.value === math.source &&
      Object.keys(attr).every((key) => ['type', 'name', 'value', 'position'].includes(key))
    );
  }
  if (opaque(payload)) {
    serializeOpaqueMdx(payload);
    return targetAstKey(payload.mdast) === targetAstKey(node);
  }
  const definition = DOMAIN_BLOCK_TAGS.get(payload.type);
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
  const expected = domainNodeFromAst({
    type: node.type,
    name: definition.tag,
    attributes,
    children: [],
  });
  const actual = domainNodeFromAst(node);
  return (
    actual !== null &&
    expected !== null &&
    jsonDataKey(actual) ===
      jsonDataKey(
        payload.type === 'source_excerpt' ? { ...expected, type: 'source_excerpt' } : expected
      )
  );
}

/** Call only after the bounded plain-JSON gate. Atomic payloads own their complete AST. */
export function isInsertedCellFragments(
  cells: unknown,
  payloads: unknown,
  inlineAtoms = false
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
        if (!binding || !payloadMatches(node, binding.payload)) return false;
        matched++;
        // Atomic payloads own their complete AST. Nested JSX is not a second native control.
        if (
          opaque(binding.payload) ||
          DOMAIN_BLOCK_TAGS.get(binding.payload.type)?.kind === 'atomic'
        )
          return true;
      } else if (binding) {
        if (node.type !== 'code' || !payloadMatches(node, binding.payload)) return false;
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
