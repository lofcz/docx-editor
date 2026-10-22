import { authorableHyperlinkTarget } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { DEFAULT_TARGET_PARSE_LIMITS, type TargetAstNode } from '../projection/parse-target';

/** Semantic intent only. Native writers retain properties that Markdown does not represent. */
export type InlineTargetMark = 'bold' | 'italic' | 'strike' | 'code';
export interface InlineTargetHyperlink {
  /** Original AST child path, not a generated native ID. Empty links retain their own entry. */
  readonly path: readonly number[];
  readonly target:
    | { readonly kind: 'external'; readonly url: string }
    | { readonly kind: 'bookmark'; readonly anchor: string };
  readonly title: string | null;
}
interface InlineTokenBase {
  readonly path: readonly number[];
  readonly marks: readonly InlineTargetMark[];
  /** The owning link's AST path. Equal URLs do not merge separate links. */
  readonly hyperlink: readonly number[] | null;
}
export type InlineTargetToken = InlineTokenBase &
  (
    | { readonly kind: 'text'; readonly value: string }
    | { readonly kind: 'atom'; readonly ast: TargetAstNode }
  );
export interface CompiledInlineTarget {
  /** Complete detached source, including unknown metadata and diagnostic positions. */
  readonly ast: TargetAstNode;
  readonly tokens: readonly InlineTargetToken[];
  readonly hyperlinks: readonly InlineTargetHyperlink[];
  /** Null means an atom needs native ownership/length resolution. Never flatten its fallback. */
  readonly plainText: string | null;
}
export type InlineTargetResult =
  | { readonly success: true; readonly target: CompiledInlineTarget }
  | NativeEditFailure;
export const INLINE_TARGET_LIMITS = DEFAULT_TARGET_PARSE_LIMITS;

const markOrder: readonly InlineTargetMark[] = ['bold', 'italic', 'strike', 'code'];
const marks = new Map<string, InlineTargetMark>([
  ['strong', 'bold'],
  ['emphasis', 'italic'],
  ['delete', 'strike'],
]);
const htmlMarks = new Map<string, InlineTargetMark>([
  ['strong', 'bold'],
  ['b', 'bold'],
  ['em', 'italic'],
  ['i', 'italic'],
  ['del', 'strike'],
  ['s', 'strike'],
  ['code', 'code'],
]);
const blocks = new Set([
  'root',
  'paragraph',
  'heading',
  'blockquote',
  'list',
  'listItem',
  'table',
  'tableRow',
  'tableCell',
  'code',
  'math',
  'thematicBreak',
  'definition',
  'footnoteDefinition',
]);
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const loneSurrogate = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const xmlText = (value: string) =>
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) && !loneSurrogate.test(value);
const modelText = (value: string) =>
  !/[\u0000-\u0008\u000B-\u001D\uFFFE\uFFFF]/.test(value) && !loneSurrogate.test(value);
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
function freeze(value: unknown): void {
  if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}
/** Extra meaning stays atomic. Locations and empty parser data are incidental. */
function standard(node: TargetAstNode, fields: readonly string[]): boolean {
  return Object.keys(node).every(
    (key) =>
      key === 'position' ||
      fields.includes(key) ||
      (key === 'data' && record(node.data) && Object.keys(node.data).length === 0)
  );
}

/**
 * Compile parsed inline data, without native IDs, tree writes, relationships, or expression evaluation.
 * Call the same function on both sides. D15 maps tokens/atoms to original canonical ownership.
 * Block metadata stays in ast. Authored NBSP stays text; source mapping owns generated task syntax.
 */
export function compileInlineTarget(input: TargetAstNode): InlineTargetResult {
  const copied = copyBoundedJson(input, INLINE_TARGET_LIMITS);
  if (!copied.ok)
    return fail(
      copied.reason === 'byte-limit' ? 'Inline target exceeds its byte limit.' : copied.reason
    );
  if (
    !record(input) ||
    Object.hasOwn(input, 'kind') ||
    typeof input.type !== 'string' ||
    !['paragraph', 'heading', 'tableCell'].includes(input.type) ||
    !Array.isArray(input.children)
  )
    return fail('Inline compilation needs a paragraph, heading, or table-cell AST.');
  const ast = copied.value as TargetAstNode,
    tokens: InlineTargetToken[] = [],
    hyperlinks: InlineTargetHyperlink[] = [];
  let hasAtom = false;
  const text = (
    path: readonly number[],
    value: string,
    active: readonly InlineTargetMark[],
    hyperlink: readonly number[] | null
  ): NativeEditFailure | null => {
    if (!modelText(value))
      return fail('Inline target text contains an invalid native model character.');
    tokens.push({ kind: 'text', path, value, marks: active, hyperlink });
    return null;
  };
  const walk = (
    node: TargetAstNode,
    path: readonly number[],
    active: readonly InlineTargetMark[],
    hyperlink: readonly number[] | null
  ): NativeEditFailure | null => {
    if (!record(node) || typeof node.type !== 'string' || !node.type || Object.hasOwn(node, 'kind'))
      return fail('Invalid inline AST node.');
    const atom = () => {
      hasAtom = true;
      tokens.push({ kind: 'atom', path, ast: node, marks: active, hyperlink });
      return null;
    };
    const children = (next = active, owner = hyperlink): NativeEditFailure | null => {
      if (!Array.isArray(node.children)) return fail('Inline wrappers need a child array.');
      for (let i = 0; i < node.children.length; i++) {
        const error = walk(node.children[i] as TargetAstNode, [...path, i], next, owner);
        if (error) return error;
      }
      return null;
    };
    const withMark = (mark: InlineTargetMark) =>
      children(markOrder.filter((item) => item === mark || active.includes(item)));
    if (node.type === 'text' || node.type === 'inlineCode') {
      if (typeof node.value !== 'string') return fail('Inline text needs a literal string value.');
      if (!standard(node, ['type', 'value'])) return atom();
      return text(
        path,
        node.value,
        node.type === 'inlineCode'
          ? markOrder.filter((mark) => mark === 'code' || active.includes(mark))
          : active,
        hyperlink
      );
    }
    if (node.type === 'break')
      return standard(node, ['type']) ? text(path, '\n', active, hyperlink) : atom();
    const mark = marks.get(node.type);
    if (mark) return standard(node, ['type', 'children']) ? withMark(mark) : atom();
    if (node.type === 'link') {
      if (!standard(node, ['type', 'url', 'title', 'children'])) return atom();
      if (hyperlink !== null)
        return fail(
          'Nested hyperlinks cannot share native inline ownership.',
          'unsupported-operation'
        );
      if (
        typeof node.url !== 'string' ||
        (node.title !== undefined && node.title !== null && typeof node.title !== 'string') ||
        (typeof node.title === 'string' && !xmlText(node.title))
      )
        return fail('Invalid inline hyperlink target or title.');
      let target: InlineTargetHyperlink['target'];
      if (node.url.startsWith('#')) {
        const anchor = node.url.slice(1);
        if (!anchor || !xmlText(anchor)) return fail('Invalid inline bookmark target.');
        target = { kind: 'bookmark', anchor };
      } else {
        const url = authorableHyperlinkTarget(node.url);
        if (url === null)
          return fail('The native engine refuses this hyperlink target.', 'unsupported-operation');
        target = { kind: 'external', url };
      }
      hyperlinks.push({ path, target, title: typeof node.title === 'string' ? node.title : null });
      return children(active, path);
    }
    if (node.type === 'mdxJsxTextElement' || node.type === 'mdxJsxFlowElement') {
      const bare =
        standard(node, ['type', 'name', 'attributes', 'children']) &&
        Array.isArray(node.attributes) &&
        node.attributes.length === 0;
      if (bare && typeof node.name === 'string') {
        const mark = htmlMarks.get(node.name);
        if (mark) return withMark(mark);
        if (node.name === 'br' && Array.isArray(node.children) && !node.children.length)
          return text(path, '\n', active, hyperlink);
      }
      return atom();
    }
    if (blocks.has(node.type))
      return fail('A block AST cannot enter an inline stream.', 'unsupported-operation');
    return atom();
  };
  for (let i = 0; i < (ast.children as readonly TargetAstNode[]).length; i++) {
    const error = walk((ast.children as readonly TargetAstNode[])[i]!, [i], [], null);
    if (error) return error;
  }
  const target: CompiledInlineTarget = {
    ast,
    tokens,
    hyperlinks,
    plainText: hasAtom
      ? null
      : tokens.map((token) => (token.kind === 'text' ? token.value : '')).join(''),
  };
  freeze(target);
  return { success: true, target };
}
