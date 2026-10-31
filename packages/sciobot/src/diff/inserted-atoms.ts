import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { MdxComponentRegistry } from '../registry/contract';
import { activeMdxRegistry } from '../registry/active';
import { MdxRefusal, astFields as fields, xmlText } from '../registry/mdx-fields';
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
        /** The registered component's atom name, such as `math`. */
        readonly atom: string;
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
function freeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}
/**
 * Plan atoms in a paragraph insertion fragment. New and retained paragraphs share the same
 * coordinate contract. This does not diff existing atom ownership or publish native edits.
 * Resources remain references: planning never fetches, decodes, embeds, or mutates package parts.
 */
export function compileInsertedAtoms(
  source: Omit<InsertedInlineAtomSource, 'pkg'>,
  paragraph: TargetAstNode,
  registry: MdxComponentRegistry = activeMdxRegistry()
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
      const inline = registry.inlineMarkdown.get(ast.type)?.inline;
      let created: BlockPayload | null;
      if (inline?.fromMarkdown) created = inline.fromMarkdown(ast, id);
      else {
        if (
          !fields(ast, ['type', 'name', 'attributes', 'children']) ||
          !Array.isArray(ast.children) ||
          ast.children.length
        )
          return fail(
            'Only complete empty registered atom syntax can enter an inline control.',
            'unsupported-operation'
          );
        created = insertedPayload(ast, id, registry);
      }
      const binding = created ? registry.components.get(created.type)?.inline : undefined;
      if (!created || !binding)
        return fail(
          `Inserted ${ast.type} atoms need their own native representation.`,
          'unsupported-operation'
        );
      const shown = binding.fallback(created);
      let payload = shown.payload;
      const fallback = shown.text,
        atom = binding.atom;
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
    return { success: true, plan: freeze({ operations, modelLength: offset - source.offset }) };
  } catch (error) {
    if (error instanceof MdxRefusal) return fail(error.message, error.code);
    return fail(error instanceof Error ? error.message : 'The inline atom plan was refused.');
  }
}

/** Validate compiled atom ownership against the immutable source package before planning writes. */
export function planInsertedAtoms(
  source: InsertedInlineAtomSource,
  paragraph: TargetAstNode,
  registry: MdxComponentRegistry = activeMdxRegistry()
): { readonly success: true; readonly plan: InsertedInlineAtomPlan } | NativeEditFailure {
  const compiled = compileInsertedAtoms(source, paragraph, registry);
  if (!compiled.success) return compiled;
  const controls = compiled.plan.operations.flatMap((op) =>
    op.kind === 'inline-control' ? [{ path: op.path, payload: op.payload, label: op.label }] : []
  );
  const problem = insertedPayloadProblem(source.pkg, controls);
  return problem ?? compiled;
}
