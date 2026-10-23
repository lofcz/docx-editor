import type { Node } from 'mdast';
import type { Plugin } from 'unified';
import type { Parent } from 'unist';
import { visit } from 'unist-util-visit';
import type { JsonValue } from '../data/json';
import type { TargetAstNode } from './parse-target';

/** Match the pinned remark-math data shape, including the source used by its static renderer. */
export function nativeMathAst(kind: 'inline' | 'block', source: string): TargetAstNode {
  const code = { type: 'text', value: source };
  return kind === 'inline'
    ? {
        type: 'inlineMath',
        value: source,
        data: {
          hName: 'code',
          hProperties: { className: ['language-math', 'math-inline'] },
          hChildren: [code],
        },
      }
    : {
        type: 'math',
        meta: null,
        value: source,
        data: {
          hName: 'pre',
          hChildren: [
            {
              type: 'element',
              tagName: 'code',
              properties: { className: ['language-math', 'math-display'] },
              children: [code],
            },
          ],
        },
      };
}
function equal(a: JsonValue, b: JsonValue): boolean {
  if (a === b) return true;
  if (
    !a ||
    !b ||
    typeof a !== 'object' ||
    typeof b !== 'object' ||
    Array.isArray(a) !== Array.isArray(b)
  )
    return false;
  const keys = Object.keys(a),
    other = Object.keys(b);
  return (
    keys.length === other.length &&
    keys.every(
      (key) =>
        Object.hasOwn(b, key) &&
        equal((a as Record<string, JsonValue>)[key]!, (b as Record<string, JsonValue>)[key]!)
    )
  );
}
/** Extension fields cannot vanish while a new equation becomes an application payload. */
export function matchesNativeMathAst(
  node: TargetAstNode,
  kind: 'inline' | 'block',
  source: string
): boolean {
  const expected = nativeMathAst(kind, source);
  if (
    node.type !== expected.type ||
    node.value !== source ||
    (node.meta !== undefined && node.meta !== null)
  )
    return false;
  if (
    Object.keys(node).some(
      (key) =>
        !['type', 'value', 'data', 'position', ...(kind === 'block' ? ['meta'] : [])].includes(key)
    )
  )
    return false;
  return (
    node.data === undefined ||
    (node.data !== null &&
      typeof node.data === 'object' &&
      !Array.isArray(node.data) &&
      !Object.keys(node.data).length) ||
    equal(node.data, expected.data!)
  );
}

/** Data-only escape hatch for sources math fences cannot preserve. Unknown fields stay opaque. */
export const remarkNativeMath: Plugin = () => (tree) => {
  visit(
    tree,
    ['mdxJsxTextElement', 'mdxJsxFlowElement'],
    (raw: Node, index: number | undefined, parent: Parent | undefined) => {
      if (!parent || typeof index !== 'number') return;
      const node = raw as unknown as TargetAstNode;
      const kind =
        node.name === 'inline_equation'
          ? 'inline'
          : node.name === 'equation' && node.type === 'mdxJsxFlowElement'
            ? 'block'
            : null;
      if (
        !kind ||
        !Array.isArray(node.children) ||
        node.children.length ||
        !Array.isArray(node.attributes) ||
        node.attributes.length !== 1
      )
        return;
      if (
        Object.keys(node).some(
          (key) => !['type', 'name', 'attributes', 'children', 'position'].includes(key)
        )
      )
        return;
      const attr = node.attributes[0];
      if (
        !attr ||
        typeof attr !== 'object' ||
        Array.isArray(attr) ||
        attr.type !== 'mdxJsxAttribute' ||
        attr.name !== 'texExpression' ||
        typeof attr.value !== 'string'
      )
        return;
      if (Object.keys(attr).some((key) => !['type', 'name', 'value', 'position'].includes(key)))
        return;
      const math = {
        ...nativeMathAst(kind, attr.value),
        ...(node.position ? { position: node.position } : {}),
      };
      parent.children[index] = (kind === 'inline' && node.type === 'mdxJsxFlowElement'
        ? {
            type: 'paragraph',
            children: [math],
            ...(node.position ? { position: node.position } : {}),
          }
        : math) as unknown as Node;
    }
  );
};
