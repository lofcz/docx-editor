import { copyBoundedJson, type JsonObject } from '../data/json';
import type { NativeEditFailure } from '../contracts/document';
import { DOMAIN_BLOCK_TAGS, type BlockTagDefinition } from '../projection/block-tags';
import type {
  ParsedTargetMarkdown,
  TargetAstNode,
  TargetAstRoot,
} from '../projection/parse-target';
import { dataFingerprint } from '../projection/protected-snapshots';
import { domainNodeFromAst } from './domain-ast';
import { targetAstKey } from './target-key';
import type {
  LogicalBlock,
  LogicalGroup,
  LogicalParagraph,
  LogicalListContext,
  LogicalRowGroup,
  LogicalBlockDocument,
  LogicalBlockResult,
} from './logical-block-types';
export type * from './logical-block-types';

export const LOGICAL_BLOCK_LIMITS = Object.freeze({
  maxBytes: 16 * 1024 * 1024,
  maxNodes: 200_000,
  maxDepth: 96,
  maxBlocks: 100_000,
});
type Definition = Pick<BlockTagDefinition, 'tag' | 'kind'>;
type Scope = { stripLists: boolean; stripPrefix: boolean };
const ordinary: Scope = { stripLists: false, stripPrefix: false };
const stripped = new Set([
  'recap_goals',
  'recap_curriculum',
  'recap_materials',
  'recap_adaptations',
]);
const blockTypes = new Set([
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
  'html',
  'definition',
  'footnoteDefinition',
  'mdxJsxFlowElement',
  'mdxFlowExpression',
]);
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const jsx = (node: TargetAstNode) =>
  node.type === 'mdxJsxFlowElement' || node.type === 'mdxJsxTextElement';
class Refusal extends Error {
  constructor(
    message: string,
    readonly code: NativeEditFailure['code'] = 'invalid-input'
  ) {
    super(message);
  }
}
const refuse = (message: string, code: NativeEditFailure['code'] = 'invalid-input'): never => {
  throw new Refusal(message, code);
};
const node = (value: unknown): TargetAstNode => {
  if (!record(value) || typeof value.type !== 'string' || !value.type)
    return refuse('Invalid logical AST node.');
  return value as TargetAstNode;
};
const children = (value: TargetAstNode): readonly TargetAstNode[] => {
  if (!Array.isArray(value.children)) return refuse('Logical AST containers need child arrays.');
  return value.children.map(node);
};
const header = (value: TargetAstNode): JsonObject =>
  Object.fromEntries(
    Object.entries(value).filter(([key]) => key !== 'children' && key !== 'position')
  );
const bare = (value: TargetAstNode, fields: readonly string[]) =>
  Object.keys(value).every(
    (key) =>
      fields.includes(key) ||
      key === 'position' ||
      (key === 'data' && record(value.data) && Object.keys(value.data).length === 0)
  );
const whitespace = (value: TargetAstNode) =>
  value.type === 'text' &&
  typeof value.value === 'string' &&
  !value.value.trim() &&
  bare(value, ['type', 'value']);
function freeze(value: unknown): void {
  if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}
function trimPrefix(ast: TargetAstNode): { ast: TargetAstNode; count: number } {
  const trim = (child: TargetAstNode): { ast: TargetAstNode; count: number } => {
    if ((child.type === 'text' || child.type === 'inlineCode') && typeof child.value === 'string') {
      const value = child.value.replace(/^\s*\d+\.\s+/, '');
      return {
        ast: value === child.value ? child : { ...child, value },
        count: child.value.length - value.length,
      };
    }
    if (
      ['strong', 'emphasis', 'delete'].includes(child.type) &&
      Array.isArray(child.children) &&
      child.children.length
    ) {
      const first = trim(node(child.children[0]));
      return first.count
        ? {
            ast: { ...child, children: [first.ast, ...child.children.slice(1)] },
            count: first.count,
          }
        : { ast: child, count: 0 };
    }
    return { ast: child, count: 0 };
  };
  const all = children(ast);
  if (!all.length) return { ast, count: 0 };
  const first = trim(all[0]!);
  return first.count
    ? { ast: { ...ast, children: [first.ast, ...all.slice(1)] }, count: first.count }
    : { ast, count: 0 };
}

/**
 * Normalize structural read data only. Apply the same transform to source and target parses.
 * Headers/ASTs retain data; keys are matching hints and cannot decide no-op edits or writability.
 * Group paths and syntax starts are not native list IDs or instructions to renumber source items.
 */
export function normalizeLogicalBlocks(
  input: Pick<ParsedTargetMarkdown, 'ast' | 'markdown'>,
  definitions: ReadonlyMap<string, Definition> = DOMAIN_BLOCK_TAGS
): LogicalBlockResult {
  try {
    if (!input || typeof input !== 'object') return refuse('Missing parsed logical target.');
    const descriptors = Object.getOwnPropertyDescriptors(input);
    for (const key of ['ast', 'markdown'])
      if (!descriptors[key] || !('value' in descriptors[key]!) || !descriptors[key]!.enumerable)
        return refuse('Logical target fields must be own data values.');
    const copy = copyBoundedJson(
      { ast: input.ast, markdown: input.markdown },
      LOGICAL_BLOCK_LIMITS
    );
    if (!copy.ok) return refuse(copy.reason);
    const parsed = copy.value as unknown as Pick<ParsedTargetMarkdown, 'ast' | 'markdown'>;
    if (
      typeof parsed.markdown !== 'string' ||
      !record(parsed.ast) ||
      parsed.ast.type !== 'root' ||
      !Array.isArray(parsed.ast.children)
    )
      return refuse('Logical targets need a parsed root and Markdown.');
    const byTag = new Map<string, { type: string; kind: Definition['kind'] }[]>();
    for (const [type, definition] of definitions) {
      if (
        !type ||
        typeof definition.tag !== 'string' ||
        !definition.tag ||
        !['container', 'atomic'].includes(definition.kind)
      )
        return refuse('Invalid logical block registry.');
      const entries = byTag.get(definition.tag) ?? [];
      entries.push({ type, kind: definition.kind });
      byTag.set(definition.tag, entries);
    }
    const policy = (ast: TargetAstNode) => {
      if (!jsx(ast) || typeof ast.name !== 'string') return null;
      const semantic = domainNodeFromAst(ast),
        definition = semantic && definitions.get(semantic.type!);
      if (definition?.tag === ast.name) return { type: semantic!.type!, kind: definition.kind };
      const candidates = byTag.get(ast.name) ?? [];
      if (!candidates.length) return null;
      if (candidates.some((candidate) => candidate.kind !== candidates[0]!.kind))
        return refuse('Ambiguous logical block registry.', 'ambiguous-target');
      return {
        type: candidates.length === 1 ? candidates[0]!.type : null,
        kind: candidates[0]!.kind,
      };
    };
    let count = 0;
    const budget = (depth: number) => {
      if (++count > LOGICAL_BLOCK_LIMITS.maxBlocks || depth > 64)
        return refuse('Logical block resource limit exceeded.');
    };
    const key = (
      kind: string,
      ast: TargetAstNode,
      descendants: readonly LogicalBlock[] = [],
      rowGroups: readonly LogicalRowGroup[] = []
    ) => {
      const props = { ...header(ast) };
      if (jsx(ast) && kind === 'container') props.type = 'mdxJsxFlowElement';
      if (['table', 'row', 'cell'].includes(kind)) {
        props.type = kind === 'table' ? 'table' : kind === 'row' ? 'tableRow' : 'tableCell';
        if (props.name === 'table' || props.name === 'tr' || props.name === 'td') delete props.name;
        if (Array.isArray(props.attributes) && !props.attributes.length) delete props.attributes;
        if (Array.isArray(props.align) && props.align.every((value) => value === null))
          delete props.align;
      }
      return `${kind}:${dataFingerprint({ own: targetAstKey(props), children: descendants.map((block) => block.key), groups: rowGroups.map((group) => targetAstKey(group.header)) })}`;
    };
    const paragraph = (
      ast: TargetAstNode,
      path: readonly number[],
      inlinePaths: readonly (readonly number[])[],
      synthetic: boolean,
      scope: Scope,
      depth: number,
      html: JsonObject | null = null
    ): LogicalParagraph => {
      budget(depth);
      const normalized = scope.stripPrefix ? trimPrefix(ast) : { ast, count: 0 };
      const meaningfulHtml =
        html &&
        Object.keys(html).some(
          (key) =>
            !['type', 'name'].includes(key) &&
            !(key === 'attributes' && Array.isArray(html.attributes) && !html.attributes.length)
        );
      return {
        kind: 'paragraph',
        path,
        list: null,
        ast: normalized.ast,
        inlinePaths,
        synthetic,
        prefixRemoved: normalized.count,
        html,
        key: `paragraph:${dataFingerprint({ body: targetAstKey(normalized.ast), html: meaningfulHtml ? targetAstKey({ ...html, type: 'mdxJsxFlowElement' }) : null })}`,
      };
    };
    const group = (
      kind: LogicalGroup['kind'],
      ast: TargetAstNode,
      path: readonly number[],
      descendants: readonly LogicalBlock[],
      domainType: string | null = null,
      rowGroups: readonly LogicalRowGroup[] = []
    ): LogicalGroup => ({
      kind,
      path,
      list: null,
      header: header(ast),
      domainType,
      children: descendants,
      rowGroups,
      key: key(kind, ast, descendants, rowGroups),
    });
    const atom = (
      ast: TargetAstNode,
      path: readonly number[],
      kind: 'atomic' | 'block',
      depth: number
    ): LogicalBlock => {
      budget(depth);
      return { kind, path, list: null, ast, key: `${kind}:${dataFingerprint(targetAstKey(ast))}` };
    };
    const html = (ast: TargetAstNode, name: string) => jsx(ast) && ast.name === name;
    const knownBlock = (ast: TargetAstNode) =>
      blockTypes.has(ast.type) ||
      (jsx(ast) && (policy(ast)?.kind === 'container' || html(ast, 'table') || html(ast, 'p')));

    const flow = (
      items: readonly TargetAstNode[],
      parent: readonly number[],
      scope: Scope,
      level: number,
      depth: number
    ): LogicalBlock[] => {
      const result: LogicalBlock[] = [],
        pending: { ast: TargetAstNode; path: readonly number[] }[] = [];
      const flush = () => {
        if (!pending.length) return;
        if (!pending.every((item) => whitespace(item.ast)))
          result.push(
            paragraph(
              { type: 'paragraph', children: pending.map((item) => item.ast) },
              parent,
              pending.map((item) => item.path),
              true,
              scope,
              depth
            )
          );
        pending.length = 0;
      };
      for (let i = 0; i < items.length; i++) {
        const ast = node(items[i]),
          path = [...parent, i];
        if (knownBlock(ast)) {
          flush();
          for (const block of convert(ast, path, scope, level, depth + 1)) result.push(block);
        } else pending.push({ ast, path });
      }
      flush();
      return result;
    };
    const list = (
      ast: TargetAstNode,
      path: readonly number[],
      scope: Scope,
      level: number,
      depth: number
    ): LogicalBlock[] => {
      budget(depth);
      if (
        typeof ast.ordered !== 'boolean' ||
        (ast.ordered && (!Number.isSafeInteger(ast.start) || Number(ast.start) < 0))
      )
        return refuse('Invalid logical list marker or start.');
      const result: LogicalBlock[] = [];
      children(ast).forEach((item, index) => {
        if (
          item.type !== 'listItem' ||
          (item.checked !== undefined && item.checked !== null && typeof item.checked !== 'boolean')
        )
          return refuse('Invalid logical list item.');
        const itemPath = [...path, index],
          blocks = flow(children(item), itemPath, scope, level + 1, depth + 1);
        if (!blocks.some((block) => block.list === null))
          blocks.unshift(
            paragraph({ type: 'paragraph', children: [] }, itemPath, [], true, scope, depth + 1)
          );
        let first = true;
        for (const block of blocks) {
          if (block.list !== null) {
            result.push(block);
            continue;
          }
          const context: LogicalListContext = {
            marker: ast.ordered ? 'ordered' : 'bullet',
            checked: typeof item.checked === 'boolean' ? item.checked : null,
            level,
            group: path,
            item: itemPath,
            role: first ? 'item' : 'continuation',
            start: ast.ordered ? Number(ast.start) : null,
            suppressed: scope.stripLists,
          };
          first = false;
          result.push({ ...block, list: context });
        }
      });
      return result;
    };
    const table = (ast: TargetAstNode, path: readonly number[], depth: number): LogicalBlock => {
      budget(depth);
      const rows: LogicalBlock[] = [],
        rowGroups: LogicalRowGroup[] = [];
      const row = (raw: TargetAstNode, rowPath: readonly number[], isHtml: boolean) => {
        budget(depth + 1);
        const cells: LogicalBlock[] = [];
        children(raw).forEach((cell, index) => {
          if (whitespace(cell)) return;
          if (
            (isHtml && !html(cell, 'td') && !html(cell, 'th')) ||
            (!isHtml && cell.type !== 'tableCell')
          )
            return refuse('Invalid logical table cell.');
          budget(depth + 2);
          const cellPath = [...rowPath, index];
          cells.push(
            group('cell', cell, cellPath, flow(children(cell), cellPath, ordinary, 0, depth + 3))
          );
        });
        rows.push(group('row', raw, rowPath, cells));
      };
      if (ast.type === 'table')
        children(ast).forEach((raw, index) => {
          if (raw.type !== 'tableRow') return refuse('Invalid logical table row.');
          row(raw, [...path, index], false);
        });
      else {
        const all = children(ast);
        const valid = (nodes: readonly TargetAstNode[], groupsAllowed: boolean): boolean =>
          nodes.every(
            (child) =>
              whitespace(child) ||
              (html(child, 'tr') &&
                children(child).every(
                  (cell) => whitespace(cell) || html(cell, 'td') || html(cell, 'th')
                )) ||
              (groupsAllowed &&
                ['thead', 'tbody', 'tfoot'].some((name) => html(child, name)) &&
                valid(children(child), false))
          );
        if (!valid(all, true)) return atom(ast, path, 'atomic', depth);
        all.forEach((raw, index) => {
          if (whitespace(raw)) return;
          const ownPath = [...path, index];
          if (html(raw, 'tr')) row(raw, ownPath, true);
          else {
            const start = rows.length;
            children(raw).forEach((child, index) => {
              if (!whitespace(child)) row(child, [...ownPath, index], true);
            });
            rowGroups.push({ path: ownPath, header: header(raw), start, end: rows.length });
          }
        });
      }
      return group('table', ast, path, rows, null, rowGroups);
    };
    const convert = (
      ast: TargetAstNode,
      path: readonly number[],
      scope: Scope,
      level: number,
      depth: number
    ): LogicalBlock[] => {
      if (depth > 64) return refuse('Logical block depth limit exceeded.');
      if (['listItem', 'tableRow', 'tableCell'].includes(ast.type))
        return refuse('An orphan structural node cannot enter a block stream.');
      if (ast.type === 'list') return list(ast, path, scope, level, depth);
      const definition = policy(ast);
      if (ast.type === 'paragraph' || ast.type === 'heading') {
        const all = children(ast),
          meaningful = all
            .map((node, index) => ({ node, index }))
            .filter(({ node }) => !whitespace(node));
        if (
          ast.type === 'paragraph' &&
          bare(ast, ['type', 'children']) &&
          meaningful.length === 1
        ) {
          const candidate = meaningful[0]!;
          if (
            jsx(candidate.node) &&
            (policy(candidate.node)?.kind === 'container' ||
              html(candidate.node, 'table') ||
              html(candidate.node, 'p'))
          )
            return convert(candidate.node, [...path, candidate.index], scope, level, depth + 1);
        }
        return [
          paragraph(
            ast,
            path,
            all.map((_, index) => [...path, index]),
            false,
            scope,
            depth
          ),
        ];
      }
      if (definition?.kind === 'container') {
        budget(depth);
        const next = {
          stripLists: stripped.has(definition.type ?? ''),
          stripPrefix: definition.type === 'recap_goals',
        };
        return [
          group(
            'container',
            ast,
            path,
            flow(children(ast), path, next, 0, depth + 1),
            definition.type
          ),
        ];
      }
      if (definition?.kind === 'atomic') {
        if (ast.name === 'source' && definition.type === 'source')
          return [
            paragraph(
              { type: 'paragraph', children: [{ ...ast, type: 'mdxJsxTextElement' }] },
              path,
              [path],
              true,
              scope,
              depth
            ),
          ];
        return [atom(ast, path, 'atomic', depth)];
      }
      if (ast.type === 'table' || html(ast, 'table')) return [table(ast, path, depth)];
      if (ast.type === 'blockquote') {
        budget(depth);
        return [group('container', ast, path, flow(children(ast), path, ordinary, 0, depth + 1))];
      }
      if (html(ast, 'p')) {
        const all = children(ast),
          normalized = { type: 'paragraph', children: all } as TargetAstNode;
        return [
          paragraph(
            normalized,
            path,
            all.map((_, index) => [...path, index]),
            true,
            scope,
            depth,
            header(ast)
          ),
        ];
      }
      return [
        atom(ast, path, jsx(ast) || ast.type === 'mdxFlowExpression' ? 'atomic' : 'block', depth),
      ];
    };
    const ast = parsed.ast as TargetAstRoot,
      blocks = flow(children(ast), [], ordinary, 0, 0);
    const document: LogicalBlockDocument = { markdown: parsed.markdown, ast, blocks };
    freeze(document);
    return { success: true, document };
  } catch (error) {
    return {
      success: false,
      code: error instanceof Refusal ? error.code : 'invalid-input',
      error: error instanceof Error ? error.message : 'Unable to normalize logical blocks.',
    };
  }
}
