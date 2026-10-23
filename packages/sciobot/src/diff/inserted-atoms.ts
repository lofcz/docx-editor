import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { createMathPayload, mathWordFallback } from '../blocks/math-codec';
import { encodeBlockPayload, type BlockPayload } from '../blocks/payload-schema';
import type { NativeEditFailure } from '../contracts/document';
import type { TargetAstNode } from '../projection/parse-target';
import {
  compileInlineTarget,
  type InlineTargetHyperlink,
  type InlineTargetMark,
} from './inline-target';
import { insertedPayload } from './inserted-blocks';
import { insertedPayloadProblem } from './inserted-payloads';
import { jsonDataKey } from './target-key';

/** Resolve new paragraph paths only after the structural executor allocates their native IDs. */
export type InlineInsertionLocation =
  | { readonly kind: 'new-paragraph'; readonly path: readonly number[] }
  | { readonly kind: 'retained-paragraph'; readonly paragraphId: string };
interface AtomPlacement {
  readonly location: InlineInsertionLocation;
  /** Original AST path inside the supplied paragraph fragment. */
  readonly path: readonly number[];
  /** UTF-16 offset after earlier text and atom fallbacks in this insertion fragment. */
  readonly offset: number;
  readonly marks: readonly InlineTargetMark[];
  readonly hyperlink: InlineTargetHyperlink | null;
}
export type InsertedInlineAtom = AtomPlacement &
  (
    | {
        readonly kind: 'inline-control';
        readonly atom: 'math' | 'mention' | 'source';
        readonly payload: BlockPayload;
        readonly label: string;
        readonly fallback: string;
      }
    | {
        readonly kind: 'drawing';
        readonly atom: 'image';
        readonly resource: {
          readonly reference: string;
          readonly alt: string;
          readonly title: string | null;
        };
        readonly modelLength: 1;
      }
  );
export interface InsertedInlineAtomPlan {
  readonly operations: readonly InsertedInlineAtom[];
  /** Model width of the fragment, including literal text and native control fallback widths. */
  readonly modelLength: number;
}
export interface InsertedInlineAtomSource {
  readonly pkg: OoxmlPackage;
  readonly location: InlineInsertionLocation;
  readonly offset: number;
  /** Called only for controls. The executor resolves image resources without persistent block IDs. */
  readonly allocateBlockId?: () => string;
}
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const xmlText = (value: unknown): value is string =>
  typeof value === 'string' &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) &&
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);
function freeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}
/** Recognized AST fields only: extension metadata must never disappear into native fallback. */
function fields(node: TargetAstNode, names: readonly string[]): boolean {
  return Object.keys(node).every(
    (key) =>
      key === 'position' ||
      names.includes(key) ||
      (key === 'data' &&
        node.data !== null &&
        typeof node.data === 'object' &&
        !Array.isArray(node.data) &&
        !Object.keys(node.data).length)
  );
}
function mathFields(node: TargetAstNode): boolean {
  if (!fields(node, ['type', 'value', 'data'])) return false;
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
/**
 * Plan atoms in a paragraph insertion fragment. New and retained paragraphs share the same
 * coordinate contract. This does not diff existing atom ownership or publish native edits.
 * Resources remain references: planning never fetches, decodes, embeds, or mutates package parts.
 */
export function planInsertedAtoms(
  source: InsertedInlineAtomSource,
  paragraph: TargetAstNode
): { readonly success: true; readonly plan: InsertedInlineAtomPlan } | NativeEditFailure {
  if (!Number.isSafeInteger(source.offset) || source.offset < 0)
    return fail('The atom insertion offset must be a nonnegative model offset.');
  const location = source.location;
  if (
    location.kind === 'new-paragraph'
      ? !location.path.length ||
        !location.path.every((step) => Number.isSafeInteger(step) && step >= 0)
      : location.kind !== 'retained-paragraph' || !location.paragraphId
  )
    return fail('The atom insertion needs a new paragraph path or retained paragraph ID.');
  const compiled = compileInlineTarget(paragraph);
  if (!compiled.success) return compiled;
  const operations: InsertedInlineAtom[] = [];
  let offset = source.offset;
  const allocate = source.allocateBlockId ?? (() => crypto.randomUUID());
  try {
    for (const token of compiled.target.tokens) {
      if (token.kind === 'text') {
        offset += token.value.length;
        continue;
      }
      const ast = token.ast;
      const placement: AtomPlacement = {
        location:
          location.kind === 'new-paragraph'
            ? { kind: location.kind, path: [...location.path] }
            : { ...location },
        path: [...token.path],
        offset,
        marks: [...token.marks],
        hyperlink: token.hyperlink
          ? (compiled.target.hyperlinks.find(
              (link) => link.path.join('/') === token.hyperlink!.join('/')
            ) ?? null)
          : null,
      };
      if (ast.type === 'image') {
        if (!fields(ast, ['type', 'url', 'alt', 'title']))
          return fail(
            'Image extension metadata needs explicit native reconciliation.',
            'unsupported-operation'
          );
        if (
          !xmlText(ast.url) ||
          !ast.url ||
          (ast.alt != null && !xmlText(ast.alt)) ||
          (ast.title != null && !xmlText(ast.title))
        )
          return fail(
            'An inline image needs a literal resource reference, alternative text, and title.'
          );
        operations.push({
          ...placement,
          kind: 'drawing',
          atom: 'image',
          modelLength: 1,
          resource: {
            reference: ast.url,
            alt: ast.alt == null ? '' : (ast.alt as string),
            title: ast.title == null ? null : (ast.title as string),
          },
        });
        offset++;
        continue;
      }
      const id = allocate();
      let payload: BlockPayload, fallback: string, atom: 'math' | 'mention' | 'source';
      if (ast.type === 'inlineMath') {
        if (!mathFields(ast))
          return fail(
            'Math extension metadata needs explicit native reconciliation.',
            'unsupported-operation'
          );
        if (!xmlText(ast.value)) return fail('Inline math needs a literal LaTeX source.');
        const math = createMathPayload(id, 'inline', ast.value);
        payload = math;
        fallback = mathWordFallback(math).inlineText;
        atom = 'math';
      } else if (ast.type === 'mention') {
        if (
          !fields(ast, ['type', 'username', 'displayText', 'children']) ||
          !Array.isArray(ast.children) ||
          ast.children.length !== 1 ||
          ast.children.some(
            (child) =>
              !child ||
              typeof child !== 'object' ||
              Array.isArray(child) ||
              child.type !== 'text' ||
              !fields(child as TargetAstNode, ['type', 'value']) ||
              !xmlText(child.value)
          )
        )
          return fail(
            'Mention extension metadata needs explicit native reconciliation.',
            'unsupported-operation'
          );
        if (
          !xmlText(ast.username) ||
          !ast.username ||
          (ast.displayText !== undefined && !xmlText(ast.displayText))
        )
          return fail('A mention needs a stored identifier and literal display text.');
        if (
          ast.displayText !== undefined &&
          (ast.children[0] as TargetAstNode).value !== ast.displayText
        )
          return fail('The mention display label and text child disagree.');
        const children = ast.children as TargetAstNode[] | undefined;
        const value =
          typeof ast.displayText === 'string'
            ? ast.displayText
            : children
                ?.map((child) =>
                  child.type === 'text' && typeof child.value === 'string' ? child.value : ''
                )
                .join('') || ast.username;
        if (!xmlText(value) || !value || /[\r\n]/.test(value))
          return fail('A mention needs a readable single-line label.');
        payload = { id, type: 'mention', version: 1, attributes: { key: ast.username, value } };
        fallback = value;
        atom = 'mention';
      } else {
        if (
          !fields(ast, ['type', 'name', 'attributes', 'children']) ||
          !Array.isArray(ast.children) ||
          ast.children.length
        )
          return fail(
            'Only a complete empty source atom can enter an inline control.',
            'unsupported-operation'
          );
        const created = insertedPayload(ast, id);
        if (!created || created.type !== 'source')
          return fail(
            `Inserted ${ast.type} atoms need their own native representation.`,
            'unsupported-operation'
          );
        payload = created;
        const attrs = payload.attributes;
        fallback =
          ([attrs.title, attrs.author, attrs.url].find(
            (value) => typeof value === 'string' && value.length > 0
          ) as string | undefined) ?? 'Source';
        if (!xmlText(fallback) || /[\r\n]/.test(fallback))
          return fail('An inline source needs a readable single-line label.');
        atom = 'source';
      }
      // Validate before the allocation escapes. Source metadata is detached from caller objects.
      payload = JSON.parse(encodeBlockPayload(payload)) as BlockPayload;
      const label = Array.from(fallback).slice(0, 128).join('');
      operations.push({ ...placement, kind: 'inline-control', atom, payload, label, fallback });
      offset += fallback.length;
    }
    if (!Number.isSafeInteger(offset))
      return fail('The atom insertion exceeds model coordinate limits.');
    const controls = operations.flatMap((op) =>
      op.kind === 'inline-control' ? [{ path: op.path, payload: op.payload, label: op.label }] : []
    );
    if (new Set(controls.map((control) => control.payload.id)).size !== controls.length)
      return fail('Inserted atom block IDs must be distinct.');
    const problem = insertedPayloadProblem(source.pkg, controls);
    if (problem) return problem;
    return { success: true, plan: freeze({ operations, modelLength: offset - source.offset }) };
  } catch (error) {
    return fail(error instanceof Error ? error.message : 'The inline atom plan was refused.');
  }
}
