import {
  createMentionPayload,
  decodeMentionPayload,
  mentionWordFallback,
} from '../blocks/mention-codec';
import { createMathPayload, decodeMathPayload, mathWordFallback } from '../blocks/math-codec';
import { encodeBlockPayload } from '../blocks/payload-schema';
import type { TargetAstNode } from '../projection/parse-target';
import { jsonDataKey } from '../diff/target-key';
import type { MdxInlineBinding } from './contract';
import { MdxRefusal, astFields, xmlText } from './mdx-fields';

function mathFields(node: TargetAstNode): boolean {
  if (!astFields(node, ['type', 'value', 'data'])) return false;
  if (
    node.data === undefined ||
    (node.data !== null &&
      typeof node.data === 'object' &&
      !Array.isArray(node.data) &&
      !Object.keys(node.data).length)
  )
    return true;
  return (
    jsonDataKey(node.data) ===
    jsonDataKey({
      hName: 'code',
      hProperties: { className: ['language-math', 'math-inline'] },
      hChildren: [{ type: 'text', value: node.value! }],
    })
  );
}

/** `$…$` inline math or its JSX element; Word shows the linear source. */
export const INLINE_EQUATION_INLINE: MdxInlineBinding = {
  atom: 'math',
  types: ['inlineMath'],
  fromMarkdown(ast, id) {
    if (!mathFields(ast))
      throw new MdxRefusal(
        'Math extension metadata needs explicit native reconciliation.',
        'unsupported-operation'
      );
    if (!xmlText(ast.value)) throw new MdxRefusal('Inline math needs a literal LaTeX source.');
    return createMathPayload(id, 'inline', ast.value);
  },
  fallback(payload) {
    const math = decodeMathPayload(encodeBlockPayload(payload));
    if (math.status !== 'supported')
      throw new MdxRefusal('An inline equation needs a valid original source.');
    return { payload: math.payload, text: mathWordFallback(math.payload).inlineText };
  },
};

/** A `@user` mention node; Word shows its display label. */
export const MENTION_INLINE: MdxInlineBinding = {
  atom: 'mention',
  types: ['mention'],
  fromMarkdown(ast, id) {
    if (
      !astFields(ast, ['type', 'username', 'displayText', 'children']) ||
      !Array.isArray(ast.children) ||
      ast.children.length !== 1 ||
      ast.children.some(
        (child) =>
          !child ||
          typeof child !== 'object' ||
          Array.isArray(child) ||
          child.type !== 'text' ||
          !astFields(child as TargetAstNode, ['type', 'value']) ||
          !xmlText(child.value)
      )
    )
      throw new MdxRefusal(
        'Mention extension metadata needs explicit native reconciliation.',
        'unsupported-operation'
      );
    if (
      !xmlText(ast.username) ||
      !ast.username ||
      (ast.displayText !== undefined && !xmlText(ast.displayText))
    )
      throw new MdxRefusal('A mention needs a stored identifier and literal display text.');
    if (
      ast.displayText !== undefined &&
      (ast.children[0] as TargetAstNode).value !== ast.displayText
    )
      throw new MdxRefusal('The mention display label and text child disagree.');
    const children = ast.children as TargetAstNode[] | undefined;
    const value =
      typeof ast.displayText === 'string'
        ? ast.displayText
        : children
            ?.map((child) =>
              child.type === 'text' && typeof child.value === 'string' ? child.value : ''
            )
            .join('') || ast.username;
    if (!xmlText(value)) throw new MdxRefusal('A mention needs a readable single-line label.');
    return createMentionPayload(id, ast.username, value);
  },
  fallback(payload) {
    const mention = decodeMentionPayload(encodeBlockPayload(payload));
    if (mention.status !== 'supported')
      throw new MdxRefusal('A mention needs a stored identifier and literal display text.');
    return { payload: mention.payload, text: mentionWordFallback(mention.payload).text };
  },
};

/** An inline citation; Word shows its title, author, or URL. */
export const SOURCE_INLINE: MdxInlineBinding = {
  atom: 'source',
  fallback(payload) {
    const attrs = payload.attributes;
    const text =
      ([attrs.title, attrs.author, attrs.url].find(
        (value) => typeof value === 'string' && value.length > 0
      ) as string | undefined) ?? 'Source';
    if (!xmlText(text) || /[\r\n]/.test(text))
      throw new MdxRefusal('An inline source needs a readable single-line label.');
    return { payload, text };
  },
};
