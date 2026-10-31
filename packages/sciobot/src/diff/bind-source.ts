import { expandLogicalSourceSpans } from './logical-source-spans';
import { findNode, paragraphOffsetIndex } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { copyBoundedJson, type JsonValue } from '../data/json';
import type { BlockTagDefinition } from '../projection/block-tags';
import { activeMdxRegistry } from '../registry/active';
import {
  parseTargetMarkdown,
  targetSourceRangeAt,
  type TargetAstNode,
} from '../projection/parse-target';
import type { ProjectedTable, ProjectedTableCell } from '../projection/tables';
import { MDX_BLOCK, MDX_INLINE } from '../projection/mdx';
import {
  indexTargetSource,
  TARGET_SOURCE_LIMITS,
  type TargetSourceInput,
  type TargetSourceOwner,
} from './target-source';
import { normalizeLogicalBlocks, type LogicalBlock } from './logical-blocks';
import { compileInlineTarget } from './inline-target';
import { nearestCellAncestor } from './cell-scope';
import { jsonDataKey } from './target-key';
import type {
  BoundLogicalSource,
  BoundSourceBlock,
  BoundSourceContribution,
  BoundSourceInline,
  BoundSourceTable,
  BoundSourceResult,
} from './bound-source-types';
export type * from './bound-source-types';

const integer = (value: unknown): value is number =>
  Number.isSafeInteger(value) && Number(value) >= 0;
const object = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
class Refusal extends Error {
  constructor(
    message: string,
    readonly code: NativeEditFailure['code'] = 'invalid-document'
  ) {
    super(message);
  }
}
const refuse = (message: string, code?: NativeEditFailure['code']): never => {
  throw new Refusal(message, code);
};
function freeze(value: unknown): void {
  if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}
const spanKey = (start: number, end: number) => `${start}:${end}`;
type ExtentEntry = { owner: TargetSourceOwner; index: number; start: number; end: number };

/** Bounded interval lookup. Paragraphs are not matched by their visible text or ordinal. */
function paragraphIntervals(owners: readonly TargetSourceOwner[]) {
  const entries: ExtentEntry[] = [];
  for (const owner of owners)
    if (owner.kind === 'paragraph' && owner.visibility === 'projected')
      owner.extents.forEach((extent, index) =>
        entries.push({ owner, index, start: extent.start, end: extent.end })
      );
  entries.sort((a, b) => a.start - b.start || a.end - b.end);
  let end = -1;
  const ends = entries.map((entry) => (end = Math.max(end, entry.end)));
  return (start: number, end: number) => {
    let low = 0,
      high = entries.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (ends[middle]! < start) low = middle + 1;
      else high = middle;
    }
    const result: ExtentEntry[] = [];
    for (let index = low; index < entries.length && entries[index]!.start <= end; index++) {
      const entry = entries[index]!;
      // An empty paragraph projects to nothing. At a block boundary it belongs to neither
      // neighbor; it stays a native sibling. It only belongs to an empty or enclosing block.
      if (
        entry.start === entry.end
          ? start === end
            ? start === entry.start
            : start < entry.start && entry.end < end
          : entry.start < end && start < entry.end
      )
        result.push(entry);
    }
    return result;
  };
}

/**
 * Bind a current native read projection to logical AST paths, without creating any edit plan.
 * Composite paragraphs, generated slots, opaque atoms, and missing AST owners stay explicit.
 * Native windows retain physical children. These bindings do not grant write authority.
 */
export function bindLogicalSource(
  input: TargetSourceInput,
  definitions: ReadonlyMap<string, Pick<BlockTagDefinition, 'tag' | 'kind'>> = activeMdxRegistry()
    .blockTags
): BoundSourceResult {
  const indexed = indexTargetSource(input);
  if (!indexed.success) return indexed;
  try {
    const copy = copyBoundedJson(input.projection, TARGET_SOURCE_LIMITS);
    if (!copy.ok) return refuse(copy.reason, 'invalid-input');
    const projection = copy.value as unknown as TargetSourceInput['projection'];
    if (!Array.isArray(projection.tables) || !Array.isArray(projection.cells))
      return refuse('Logical binding needs composed native table ownership.', 'invalid-input');
    for (const block of projection.blocks) {
      if ([MDX_BLOCK, MDX_INLINE].includes(block.type)) continue;
      const definition = definitions.get(block.type);
      if (definition?.tag !== block.tag || definition.kind !== block.kind)
        return refuse('The source registry differs from its native projection.', 'invalid-input');
    }
    const parsed = parseTargetMarkdown(projection.text);
    if (!parsed.success) return parsed;
    const normalized = normalizeLogicalBlocks(parsed.target, definitions);
    if (!normalized.success) return normalized;
    const expandedSource = expandLogicalSourceSpans(normalized.document.blocks);
    const index = indexed.index,
      logical = { ...normalized.document, blocks: expandedSource.blocks },
      part = input.pkg.parts.get(index.partName)!,
      owners = new Map(index.owners.map((owner) => [owner.id, owner]));
    const tables = new Map<string, ProjectedTable>(),
      tableRows = new Map<string, ReadonlySet<string>>(),
      cells = new Map<string, ProjectedTableCell>(),
      rows = new Map<string, ProjectedTableCell[]>(),
      paragraphCells = new Map<string, string>();
    for (const table of projection.tables) {
      if (
        !object(table) ||
        typeof table.id !== 'string' ||
        tables.has(table.id) ||
        (table.parentCellId !== null && typeof table.parentCellId !== 'string') ||
        !Array.isArray(table.rowIds) ||
        table.rowIds.some((id) => typeof id !== 'string') ||
        new Set(table.rowIds).size !== table.rowIds.length ||
        findNode(part, table.id)?.kind !== 'table' ||
        (nearestCellAncestor(part, table.id, 'tableCell')?.id ?? null) !== table.parentCellId
      )
        return refuse('Invalid or conflicting native table ownership.');
      for (const rowId of table.rowIds)
        if (
          findNode(part, rowId)?.kind !== 'tableRow' ||
          nearestCellAncestor(part, rowId, 'table')?.id !== table.id
        )
          return refuse('A projected row has no native table owner.');
      tables.set(table.id, table as unknown as ProjectedTable);
      tableRows.set(table.id, new Set(table.rowIds));
    }
    for (const cell of projection.cells) {
      const table = object(cell) && typeof cell.tableId === 'string' && tables.get(cell.tableId);
      if (
        !table ||
        typeof cell.id !== 'string' ||
        cells.has(cell.id) ||
        typeof cell.rowId !== 'string' ||
        !tableRows.get(table.id)!.has(cell.rowId) ||
        !integer(cell.rowIndex) ||
        !integer(cell.gridColumn) ||
        !integer(cell.gridSpan) ||
        cell.gridSpan < 1 ||
        !Number.isSafeInteger(cell.gridColumn + cell.gridSpan) ||
        typeof cell.verticalContinuation !== 'boolean' ||
        !Array.isArray(cell.paragraphIds) ||
        !Array.isArray(cell.sources) ||
        findNode(part, cell.id)?.kind !== 'tableCell' ||
        nearestCellAncestor(part, cell.id, 'tableRow')?.id !== cell.rowId ||
        nearestCellAncestor(part, cell.id, 'table')?.id !== cell.tableId
      )
        return refuse('Invalid or conflicting native cell ownership.');
      for (const id of cell.paragraphIds) {
        if (
          typeof id !== 'string' ||
          paragraphCells.has(id) ||
          findNode(part, id)?.kind !== 'paragraph' ||
          nearestCellAncestor(part, id, 'tableCell')?.id !== cell.id
        )
          return refuse('A projected paragraph has conflicting nearest cell ownership.');
        paragraphCells.set(id, cell.id);
      }
      cells.set(cell.id, cell as unknown as ProjectedTableCell);
      const siblings = rows.get(cell.rowId) ?? [];
      siblings.push(cell as unknown as ProjectedTableCell);
      rows.set(cell.rowId, siblings);
    }
    for (const siblings of rows.values()) {
      siblings.sort((a, b) => a.gridColumn - b.gridColumn);
      let end = 0;
      for (const cell of siblings) {
        if (cell.gridColumn < end)
          return refuse('Native projected grid intervals overlap.', 'ambiguous-target');
        end = cell.gridColumn + cell.gridSpan;
      }
    }
    const sourcesByCell = new Map<string, (typeof projection.sourceMap.sources)[number][]>();
    for (const slice of projection.sourceMap.sources) {
      if (slice.sourceScope !== 'body') continue;
      const cellId = paragraphCells.get(slice.paragraphId);
      if (!cellId) continue;
      const list = sourcesByCell.get(cellId) ?? [];
      list.push(slice);
      sourcesByCell.set(cellId, list);
    }
    for (const cell of cells.values())
      if (
        jsonDataKey(cell.sources as unknown as JsonValue) !==
        jsonDataKey((sourcesByCell.get(cell.id) ?? []) as unknown as JsonValue)
      )
        return refuse('Cell source slices differ from the composed projection.');

    const exactOwners = new Map<string, ExtentEntry[]>(),
      inlineByParagraph = new Map<string, TargetSourceOwner[]>();
    for (const owner of index.owners)
      if (owner.block && owner.paragraphId && owner.visibility === 'projected') {
        const list = inlineByParagraph.get(owner.paragraphId) ?? [];
        list.push(owner);
        inlineByParagraph.set(owner.paragraphId, list);
      }
    for (const owner of index.owners)
      if (owner.visibility === 'projected' && owner.kind !== 'paragraph')
        owner.extents.forEach((extent, extentIndex) => {
          const key = spanKey(extent.start, extent.end),
            list = exactOwners.get(key) ?? [];
          list.push({ owner, index: extentIndex, start: extent.start, end: extent.end });
          exactOwners.set(key, list);
        });
    const queryParagraphs = paragraphIntervals(index.owners);
    const originalNode = (path: readonly number[]): TargetAstNode | null => {
      let node: TargetAstNode = logical.ast;
      for (const index of path) {
        if (!Array.isArray(node.children) || !object(node.children[index])) return null;
        node = node.children[index] as TargetAstNode;
      }
      return node;
    };
    const position = (node: TargetAstNode | null): { start: number; end: number } | null => {
      const p = node?.position;
      if (
        !object(p) ||
        !object(p.start) ||
        !object(p.end) ||
        !integer(p.start.offset) ||
        !integer(p.end.offset) ||
        p.end.offset < p.start.offset
      )
        return null;
      return { start: p.start.offset, end: p.end.offset };
    };
    const sourceSpan = (block: LogicalBlock) => {
      if (expandedSource.virtual.has(block)) return null;
      let extent: { start: number; end: number } | null;
      if (block.kind === 'paragraph' && block.synthetic && !block.html) {
        const first = block.inlinePaths[0],
          last = block.inlinePaths.at(-1);
        const from = first && position(originalNode(first)),
          to = last && position(originalNode(last));
        extent = from && to ? { start: from.start, end: to.end } : null;
      } else extent = position(originalNode(block.path));
      return extent ? targetSourceRangeAt(parsed.target, extent.start, extent.end) : null;
    };
    type MutableBinding = Omit<BoundSourceBlock, 'structuralOwnerId'> & {
      structuralOwnerId: string | null;
    };
    const bindings: MutableBinding[] = [];
    const walk = (
      blocks: readonly LogicalBlock[],
      parent: readonly number[],
      context: BoundSourceTable | null,
      depth: number
    ) => {
      if (depth > 64) return refuse('Logical binding depth limit exceeded.', 'invalid-input');
      blocks.forEach((block, at) => {
        if (bindings.length >= TARGET_SOURCE_LIMITS.maxNodes)
          return refuse('Logical binding count limit exceeded.', 'invalid-input');
        const logicalPath = [...parent, at],
          span = sourceSpan(block),
          matches = span ? (exactOwners.get(spanKey(span.start, span.end)) ?? []) : [];
        if (matches.length > 1)
          return refuse('A logical source span has several native owners.', 'ambiguous-target');
        const exact = matches[0]?.owner ?? null;
        if (exact && context?.cellId && exact.cell?.cellId !== context.cellId)
          return refuse('A logical construct conflicts with its nearest native cell.');
        let table = context;
        if (
          exact?.kind === 'table' ||
          (block.kind === 'table' && exact?.block?.kind !== 'atomic')
        ) {
          if (!exact || exact.kind !== 'table' || !tables.has(exact.id))
            return refuse('A logical table has no exact native table owner.');
          table = {
            tableId: exact.id,
            rowId: null,
            cellId: null,
            gridColumn: null,
            gridSpan: null,
            role: 'table',
          };
          if (
            block.kind === 'table' &&
            block.children.length !== tables.get(exact.id)!.rowIds.length
          )
            return refuse('Logical and native row counts differ.');
        } else if (block.kind === 'row') {
          if (!context) return refuse('A logical row has no bound table.');
          const rowId = tables.get(context.tableId)?.rowIds[at];
          if (!rowId) return refuse('A logical row has no native row identity.');
          table = {
            ...context,
            rowId,
            cellId: null,
            gridColumn: null,
            gridSpan: null,
            role: 'row',
          };
        } else if (block.kind === 'cell') {
          if (!context?.rowId) return refuse('A logical cell has no bound row.');
          const physical = (rows.get(context.rowId) ?? []).find(
            (cell) => cell.gridColumn <= at && at < cell.gridColumn + cell.gridSpan
          );
          const role = !physical
            ? 'missing-slot'
            : physical.gridColumn !== at
              ? 'horizontal-placeholder'
              : physical.verticalContinuation
                ? 'vertical-placeholder'
                : 'cell';
          if (
            expandedSource.virtual.has(block) &&
            !['horizontal-placeholder', 'vertical-placeholder'].includes(role)
          )
            return refuse('A source HTML span conflicts with the canonical native grid.');
          table = {
            ...context,
            cellId: physical?.id ?? null,
            gridColumn: at,
            gridSpan: physical?.gridSpan ?? null,
            role,
          };
        }
        const placeholder =
          table &&
          ['horizontal-placeholder', 'vertical-placeholder', 'missing-slot'].includes(table.role);
        const contributions: BoundSourceContribution[] = [];
        const atom = exact?.block?.kind === 'atomic' ? exact : null;
        if (exact && !exact.paragraphId)
          contributions.push({
            ownerId: exact.id,
            extentIndex: matches[0]!.index,
            start: span!.start,
            end: span!.end,
          });
        else if (
          span &&
          !placeholder &&
          (!['row', 'cell', 'container'].includes(block.kind) || !!atom)
        ) {
          for (const entry of queryParagraphs(span.start, span.end)) {
            if (table?.cellId && entry.owner.cell?.cellId !== table.cellId) {
              if (entry.owner.cell?.tableId === table.tableId)
                return refuse('Logical text conflicts with its projected physical cell.');
              continue;
            }
            contributions.push({
              ownerId: entry.owner.id,
              extentIndex: entry.index,
              start: Math.max(span.start, entry.start),
              end: Math.min(span.end, entry.end),
            });
          }
        }
        const inline: BoundSourceInline[] = [];
        const paragraphIds = new Set(
          contributions
            .map((entry) => owners.get(entry.ownerId)?.paragraphId)
            .filter((id): id is string => !!id)
        );
        for (const owner of [...paragraphIds].flatMap((id) => inlineByParagraph.get(id) ?? []))
          if (
            span &&
            owner.block &&
            owner.paragraphId &&
            paragraphIds.has(owner.paragraphId) &&
            owner.visibility === 'projected'
          ) {
            const extents = owner.extents.filter(
              (extent) => span.start <= extent.start && extent.end <= span.end
            );
            if (!extents.length) continue;
            const paragraph = findNode(part, owner.paragraphId),
              control = findNode(part, owner.id);
            const native =
              paragraph?.kind === 'paragraph' &&
              control &&
              paragraphOffsetIndex(paragraph).spanOf(control);
            if (!native) return refuse('An inline block has no native paragraph range.');
            extents.forEach((extent) =>
              inline.push({
                controlId: owner.id,
                blockId: owner.block!.id,
                start: extent.start,
                end: extent.end,
                native: { start: native.start, end: native.end },
              })
            );
          }
        let generatedTask: BoundSourceBlock['generatedTask'] = null;
        const ids = [...new Set(contributions.map((entry) => entry.ownerId))];
        if (block.kind === 'paragraph' && ids.length === 1) {
          const task = owners.get(ids[0]!)?.checklist;
          if (task?.empty) {
            const compiled = compileInlineTarget(block.ast);
            if (
              !span ||
              span.start !== task.contentStart ||
              span.end !== task.contentStart + 1 ||
              projection.text.slice(span.start, span.end) !== '\u00a0' ||
              !compiled.success ||
              compiled.target.plainText !== '\u00a0' ||
              block.list?.checked !== task.checked
            )
              return refuse('Generated empty-task syntax differs from its source binding.');
            generatedTask = {
              paragraphId: task.paragraphId,
              controlId: task.controlId,
              start: span.start,
              end: span.end,
              nativeInsertionOffset: task.native.end,
            };
          }
        }
        bindings.push({
          logicalPath,
          astPath: block.path,
          source: span,
          contributions,
          structuralOwnerId: null,
          table,
          inline,
          atomicControlId: atom?.id ?? null,
          generatedTask,
        });
        if ('children' in block && !atom && !placeholder)
          walk(block.children, logicalPath, table, depth + 1);
        if (placeholder && block.kind === 'cell' && block.children.length)
          return refuse('A generated table slot contains unexpected logical content.');
      });
    };
    walk(logical.blocks, [], null, 0);
    const claims = new Map<string, MutableBinding[]>(),
      represented = new Set<string>();
    for (const binding of bindings) {
      for (const id of new Set(binding.contributions.map((entry) => entry.ownerId))) {
        const list = claims.get(id) ?? [];
        list.push(binding);
        claims.set(id, list);
        represented.add(id);
      }
      binding.inline.forEach((entry) => represented.add(entry.controlId));
    }
    const windowIds = new Set(index.windows.flatMap((window) => window.ids));
    for (const [id, list] of claims)
      if (
        list.length === 1 &&
        new Set(list[0]!.contributions.map((entry) => entry.ownerId)).size === 1 &&
        windowIds.has(id)
      )
        list[0]!.structuralOwnerId = id;
    const source: BoundLogicalSource = {
      index,
      parsed: {
        sourceMarkdown: parsed.target.sourceMarkdown,
        markdown: parsed.target.markdown,
        normalization: parsed.target.normalization,
      },
      logical,
      bindings,
      windows: index.windows.map((window) => ({
        ...window,
        members: window.ids.map((id) => ({
          ownerId: id,
          logicalPaths: (claims.get(id) ?? []).map((binding) => binding.logicalPath),
        })),
      })),
      unrepresented: index.owners
        .filter((owner) => !represented.has(owner.id))
        .map((owner) => ({
          ownerId: owner.id,
          reason:
            owner.visibility === 'projected'
              ? owner.extents.every((extent) => extent.start === extent.end)
                ? 'empty-extent'
                : 'not-parsed'
              : owner.visibility,
        })),
    };
    freeze(source);
    return { success: true, source };
  } catch (error) {
    return {
      success: false,
      code: error instanceof Refusal ? error.code : 'invalid-input',
      error: error instanceof Error ? error.message : 'Unable to bind logical source.',
    };
  }
}
