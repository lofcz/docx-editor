import { activeMdxRegistry } from '../registry/active';
import type { MarkdownLayoutOptions } from '@docx-editor.dev/docx-to-markdown';
import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { Root } from 'mdast';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import remarkMdx from 'remark-mdx';
import remarkParse from 'remark-parse';
import remarkStringify from 'remark-stringify';
import { unified } from 'unified';
import {
  encodeBlockPayload,
  type BlockPayload,
  type JsonObject,
  type JsonValue,
} from '../blocks/payload-schema';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { type BlockTagDefinition, type SerializedBlockTag } from './block-tags';
import { projectBlocks, type BlockProjection } from './blocks';
import { projectTableOwnership, type TableProjection } from './tables';

export { CUSTOM_ALLOWED_TAGS, CUSTOM_DATA_ONLY_TAGS, CUSTOM_LITERAL_TAGS } from './mdx-tags';

export const MDX_INLINE = 'shared_mdx_inline';
export const MDX_BLOCK = 'shared_mdx_block';

/** Persist the complete JSON tree, including positions and unknown metadata. */
export interface OpaqueMdxAst extends JsonObject {
  readonly type: 'mdxJsxFlowElement' | 'mdxJsxTextElement';
  readonly name: string | null;
  readonly attributes: readonly JsonObject[];
  readonly children: readonly JsonObject[];
}

export interface SerializedOpaqueMdx extends SerializedBlockTag {
  readonly ast: OpaqueMdxAst;
  readonly inline: boolean;
}

export interface OpaqueMdxSpan {
  readonly blockId: string;
  readonly controlId: string;
  readonly start: number;
  readonly end: number;
  readonly inline: boolean;
  readonly ast: OpaqueMdxAst;
}

export interface MdxProjection
  extends BlockProjection, Pick<TableProjection, 'tables' | 'cells' | 'boundaries'> {
  readonly opaqueMdx: readonly OpaqueMdxSpan[];
}

// These are the application's shared MDX plugins. This path never compiles or evaluates JavaScript.
const processor = unified()
  .use(remarkParse)
  .use(remarkMath)
  .use(remarkGfm)
  .use(remarkMdx)
  .use(remarkStringify);

function freezeJson(value: JsonValue): void {
  if (value === null || typeof value !== 'object') return;
  for (const child of Object.values(value)) freezeJson(child);
  Object.freeze(value);
}

function jsonObject(value: JsonValue | undefined): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function storedAst(payload: BlockPayload): OpaqueMdxAst {
  // Validate before JSON.stringify can invoke getters or discard unsupported values.
  // JSON.parse creates own data properties; it never merges prototype-shaped keys.
  const clone = JSON.parse(encodeBlockPayload(payload)) as BlockPayload;
  if (clone.type !== MDX_INLINE && clone.type !== MDX_BLOCK)
    throw new TypeError('Expected a shared MDX payload.');
  const ast = clone.mdast;
  if (
    !jsonObject(ast) ||
    !Object.hasOwn(ast, 'type') ||
    ast.type !== (clone.type === MDX_INLINE ? 'mdxJsxTextElement' : 'mdxJsxFlowElement') ||
    !Object.hasOwn(ast, 'name') ||
    (ast.name !== null &&
      (typeof ast.name !== 'string' ||
        !/^[\p{ID_Start}_$][\p{ID_Continue}$_.:\u200C\u200D-]*$/u.test(ast.name))) ||
    !Array.isArray(ast.attributes) ||
    !Array.isArray(ast.children)
  )
    throw new TypeError('Invalid opaque MDX root or inline identity.');
  freezeJson(ast);
  return ast as OpaqueMdxAst;
}

/** Serialize stored data with remark. Expressions remain source text and never execute. */
export function serializeOpaqueMdx(payload: BlockPayload): SerializedOpaqueMdx {
  const ast = storedAst(payload);
  const inline = payload.type === MDX_INLINE;
  const root = {
    type: 'root',
    children: inline ? [{ type: 'paragraph', children: [ast] }] : [ast],
  } as unknown as Root;
  const text = processor.stringify(root).replace(/\n+$/, '');
  // Malformed stored expression syntax must fail before becoming an AI edit target.
  processor.parse(text);
  return Object.freeze({
    ast,
    inline,
    tag: ast.name ?? 'fragment',
    text,
    properties: Object.freeze([
      Object.freeze({
        name: 'mdast',
        property: 'mdast',
        path: Object.freeze(['mdast']),
        start: 0,
        end: text.length,
        valueStart: 0,
        valueEnd: text.length,
      }),
    ]),
  });
}

/** Add complete shared MDX trees to the native/domain projection without changing Word XML. */
export function projectMdx(
  pkg: OoxmlPackage,
  layout: ExportSemanticLayout,
  revision: RevisionToken,
  definitions: ReadonlyMap<string, BlockTagDefinition> = activeMdxRegistry().blockTags,
  options: Pick<MarkdownLayoutOptions, 'tableSyntax'> = {}
): { readonly success: true; readonly projection: MdxProjection } | NativeEditFailure {
  const serialized = new Map<string, SerializedOpaqueMdx>();
  const registry = new Map(definitions);
  for (const type of [MDX_INLINE, MDX_BLOCK])
    registry.set(type, {
      tag: 'opaque',
      kind: 'atomic',
      attributes: () => [],
      serialize(payload) {
        const value = serializeOpaqueMdx(payload);
        serialized.set(payload.id, value);
        return value;
      },
    });
  const result = projectBlocks(pkg, layout, revision, registry, options);
  if (!result.success) return result;
  const ownership = projectTableOwnership(layout, result.projection);
  if (!ownership.success) return ownership;
  const opaqueMdx: OpaqueMdxSpan[] = [];
  for (const block of result.projection.blocks) {
    const value = serialized.get(block.id);
    if (value)
      opaqueMdx.push(
        Object.freeze({
          blockId: block.id,
          controlId: block.controlId,
          start: block.start,
          end: block.end,
          inline: value.inline,
          ast: value.ast,
        })
      );
  }
  return {
    success: true,
    projection: Object.freeze({
      ...result.projection,
      tables: ownership.projection.tables,
      cells: ownership.projection.cells,
      boundaries: ownership.projection.boundaries,
      opaqueMdx: Object.freeze(opaqueMdx),
    }),
  };
}
