import { matchesNativeMathAst } from '../projection/native-math';
import { createMathPayload, decodeMathPayload } from '../blocks/math-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { TargetAstNode } from '../projection/parse-target';
import type { MdxMarkdownBinding } from './contract';

const jsx = (node: TargetAstNode) =>
  node.type === 'mdxJsxFlowElement' || node.type === 'mdxJsxTextElement';

/** A fenced code block: its language and source are the whole payload. */
export const CODE_BLOCK_MARKDOWN: MdxMarkdownBinding = {
  types: ['code'],
  fromMarkdown(ast, id) {
    if (
      typeof ast.value !== 'string' ||
      (ast.lang !== null && typeof ast.lang !== 'string') ||
      (typeof ast.lang === 'string' && ast.lang.length > 128) ||
      (ast.meta !== null && ast.meta !== undefined)
    )
      return null;
    return {
      version: 1,
      id,
      type: 'code_block',
      attributes: { language: ast.lang ?? '', value: ast.value },
    };
  },
  matches: (node, payload) =>
    node.type === 'code' &&
    typeof node.value === 'string' &&
    node.value === payload.attributes.value &&
    (node.lang ?? '') === payload.attributes.language,
};

/** An equation reads from `$$`/`$` math or from its JSX element with one `texExpression`. */
function equation(kind: 'block' | 'inline'): MdxMarkdownBinding {
  return {
    types: [kind === 'block' ? 'math' : 'inlineMath'],
    fromMarkdown(ast, id): BlockPayload | null {
      if (typeof ast.value !== 'string' || !matchesNativeMathAst(ast, kind, ast.value)) return null;
      try {
        return createMathPayload(id, kind, ast.value);
      } catch {
        return null;
      }
    },
    matches(node, payload) {
      const math = decodeMathPayload(encodeBlockPayload(payload));
      if (math.status !== 'supported') return false;
      if (node.type === 'math' || node.type === 'inlineMath')
        return matchesNativeMathAst(node, math.kind, math.source);
      if (
        !jsx(node) ||
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
    },
  };
}

export const BLOCK_EQUATION_MARKDOWN = equation('block');
export const INLINE_EQUATION_MARKDOWN = equation('inline');
