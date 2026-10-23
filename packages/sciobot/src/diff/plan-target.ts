import { atomicTargetsPairable, planTargetAtomic } from './target-atomic';
import {
  logicalBlockAst as astOf,
  documentQuoteStyle,
  insertedBlockProblem,
} from './inserted-fragment-policy';
import { findNode } from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import type { NativeEditFailure } from '../contracts/document';
import type { JsonObject } from '../data/json';
import { DOMAIN_BLOCK_TAGS, type BlockTagDefinition } from '../projection/block-tags';
import { parseTargetMarkdown } from '../projection/parse-target';
import { bindLogicalSource, type BoundSourceBlock } from './bind-source';
import type { TargetSourceInput } from './target-source';
import { normalizeLogicalBlocks, type LogicalBlock } from './logical-blocks';
import { compileInlineTarget } from './inline-target';
import { planTargetParagraph, targetParagraphHtmlKey } from './target-inline';
import { planTargetTable } from './target-table';
import { prepareTargetCellParagraphs } from './target-cell-paragraphs';
import { planDomainProperties } from './domain-properties';
import { domainNodeFromAst } from './domain-ast';
import { sameContainerIdentity } from './domain-strategies';
import { planListMembership } from './lists';
import { planInsertedBlocks, type NewBlockPayload } from './inserted-blocks';
import { resolveTargetLists } from './target-lists';
import { planSiblingBlocks, siblingContentPairable } from './sibling-blocks';
import { createEditPlan, type CanonicalEdit, type CanonicalEditPlan } from './plan';
import { jsonDataKey, targetAstKey } from './target-key';
import {
  bindQuoteParagraphStream,
  quoteParagraphStream,
  quoteSiblingInsertion,
} from './target-quotes';
import { planParagraphFormat } from './paragraph-format';

export interface TargetDocumentOptions {
  readonly definitions?: ReadonlyMap<string, Pick<BlockTagDefinition, 'tag' | 'kind'>>;
  readonly headingStyle?: (depth: number) => string | null;
  /** Paragraph style for new block quotes; defaults to the document's own Quote style. */
  readonly quoteStyle?: () => string | null;
  /** Persistent IDs for new registered blocks; defaults to random UUIDs. */
  readonly allocateBlockId?: () => string;
  /** Trusted host data for new atomic React blocks. AST expressions remain inert stored text. */
  readonly newBlockPayload?: NewBlockPayload;
}
export interface TargetDocumentPair {
  readonly id: string;
  readonly beforePath: readonly number[];
  readonly afterPath: readonly number[];
}
export type TargetDocumentPlanResult =
  | {
      readonly success: true;
      readonly plan: CanonicalEditPlan;
      readonly pairs: readonly TargetDocumentPair[];
    }
  | NativeEditFailure;
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
const listKey = (block: LogicalBlock) =>
  block.list
    ? jsonDataKey(
        Object.fromEntries(
          Object.entries(block.list).filter(([name]) => !['group', 'item'].includes(name))
        )
      )
    : 'null';
const deepSame = (before: LogicalBlock, after: LogicalBlock): boolean => {
  if (
    before.kind !== after.kind ||
    targetAstKey(astOf(before)) !== targetAstKey(astOf(after)) ||
    listKey(before) !== listKey(after)
  )
    return false;
  if (before.kind === 'paragraph' && after.kind === 'paragraph')
    return targetParagraphHtmlKey(before) === targetParagraphHtmlKey(after);
  if ('children' in before && 'children' in after) {
    const groups = (block: typeof before) =>
      jsonDataKey(
        block.rowGroups.map((group) => ({
          start: group.start,
          end: group.end,
          header: targetAstKey(group.header),
        }))
      );
    return (
      groups(before) === groups(after) &&
      before.children.length === after.children.length &&
      before.children.every((child, index) => deepSame(child, after.children[index]!))
    );
  }
  return true;
};

/** Longest common subsequence of block content keys, as matched index pairs. */
function keyAlignment(before: readonly string[], after: readonly string[]): [number, number][] {
  const n = before.length,
    m = after.length,
    table = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--)
      table[i]![j] =
        before[i] === after[j]
          ? table[i + 1]![j + 1]! + 1
          : Math.max(table[i + 1]![j]!, table[i]![j + 1]!);
  const out: [number, number][] = [];
  for (let i = 0, j = 0; i < n && j < m; ) {
    if (before[i] === after[j]) out.push([i++, j++]);
    else if (table[i + 1]![j]! >= table[i]![j + 1]!) i++;
    else j++;
  }
  return out;
}

/**
 * Compose complete-target intentions over one immutable native source. Never replace its story tree.
 * D15 remains in progress: unsupported dispatch paths refuse the entire request without partial plans.
 * Direct/review executors still own protection, live revisions, resource allocation, and application.
 */
export function planTargetDocument(
  input: TargetSourceInput,
  targetMarkdown: string,
  options: TargetDocumentOptions = {}
): TargetDocumentPlanResult {
  const definitions = options.definitions ?? DOMAIN_BLOCK_TAGS;
  const bound = bindLogicalSource(input, definitions);
  if (!bound.success) return bound;
  const parsed = parseTargetMarkdown(targetMarkdown);
  if (!parsed.success) return parsed;
  const normalized = normalizeLogicalBlocks(parsed.target, definitions);
  if (!normalized.success) return normalized;
  const beforeQuotes = bindQuoteParagraphStream(bound.source),
    afterQuotes = quoteParagraphStream(normalized.document);
  const cells = prepareTargetCellParagraphs(input, beforeQuotes.source, afterQuotes.document);
  if (!cells.success) return cells;
  const source = cells.source,
    target = cells.target,
    revision = source.index.revision,
    part = input.pkg.parts.get(source.index.partName)!,
    blockIndex = buildBlockIndex(input.pkg),
    bindings = new Map(source.bindings.map((binding) => [binding.logicalPath.join('/'), binding])),
    byControl = new Map([...blockIndex.byId.values()].map((block) => [block.controlId, block])),
    operations: CanonicalEdit[] = [],
    pairs: TargetDocumentPair[] = [];
  if (targetAstKey(source.logical.ast) === targetAstKey(target.ast)) {
    const result = createEditPlan(revision, []);
    return result.success ? { ...result, pairs: Object.freeze([]) } : result;
  }
  let visits = 0;
  const stream = (
    before: readonly LogicalBlock[],
    after: readonly LogicalBlock[],
    beforePath: readonly number[],
    afterPath: readonly number[],
    parentId: string,
    offset: { readonly before: number; readonly after: number } = { before: 0, after: 0 },
    forced?: (typeof source.windows)[number]
  ): NativeEditFailure | null => {
    if (++visits > 100_000 || beforePath.length > 64)
      return fail('Complete-target planning exceeds its recursion limit.');
    const mapped = before.map((block, index) => ({
      block,
      binding: bindings.get([...beforePath, index + offset.before].join('/')),
      ast: astOf(block),
    }));
    if (mapped.some((entry) => !entry.binding?.structuralOwnerId))
      return fail('This logical stream needs composite or unrepresented native owner planning.');
    const ids = mapped.map((entry) => entry.binding!.structuralOwnerId!);
    const windows = forced
      ? [forced]
      : source.windows.filter(
          (window) =>
            window.parentId === parentId &&
            window.ids.length === ids.length &&
            window.ids.every((id, index) => id === ids[index])
        );
    if (windows.length !== 1) {
      // Native barriers (opaque blocks, unprojected nodes) split one logical stream into several
      // physical sibling windows. Plan each window separately and leave every barrier in place.
      const candidates = source.windows.filter(
        (window) => window.parentId === parentId && window.ids.length
      );
      const cover: { window: (typeof source.windows)[number]; start: number }[] = [];
      for (let at = 0; at < ids.length; ) {
        const window = candidates.find((item) =>
          item.ids.every((id, index) => ids[at + index] === id)
        );
        if (!window)
          return fail(
            'The complete target needs separate planning across native sibling barriers.'
          );
        cover.push({ window, start: at });
        at += window.ids.length;
      }
      if (cover.length < 2 || before.length * after.length > 4_000_000)
        return fail('The complete target needs separate planning across native sibling barriers.');
      const matches = keyAlignment(
        before.map((block) => block.key),
        after.map((block) => block.key)
      );
      const splits = [0];
      for (let k = 1; k < cover.length; k++) {
        const boundary = cover[k]!.start,
          last = matches.filter(([i]) => i < boundary).at(-1)?.[1],
          first = matches.find(([i]) => i >= boundary)?.[1];
        // Unmatched new blocks between two windows join the earlier window, before the barrier.
        const split = first ?? (last === undefined ? splits.at(-1)! : last + 1);
        splits.push(Math.max(split, splits.at(-1)!));
      }
      splits.push(after.length);
      for (const [k, entry] of cover.entries()) {
        const end = cover[k + 1]?.start ?? before.length;
        const error = stream(
          before.slice(entry.start, end),
          after.slice(splits[k]!, splits[k + 1]!),
          beforePath,
          afterPath,
          parentId,
          { before: offset.before + entry.start, after: offset.after + splits[k]! },
          entry.window
        );
        if (error) return error;
      }
      return null;
    }
    const window = windows[0]!,
      targets = after.map(astOf),
      oldMap = new Map(mapped.map((entry) => [entry.ast, entry])),
      newMap = new Map(targets.map((ast, index) => [ast, after[index]!]));
    const result = planSiblingBlocks(
      {
        revision,
        part,
        parentId,
        startIndex: window.startIndex,
        blocks: mapped.map((entry) => ({ id: entry.binding!.structuralOwnerId!, ast: entry.ast })),
      },
      targets,
      {
        key: (ast) => oldMap.get(ast)?.block.key ?? newMap.get(ast)!.key,
        pairable(old, next) {
          const entry = oldMap.get(old)!,
            afterBlock = newMap.get(next)!,
            native = byControl.get(entry.binding!.structuralOwnerId!);
          if (native) {
            const a = domainNodeFromAst(old),
              b = domainNodeFromAst(next);
            if (entry.binding!.atomicControlId)
              return atomicTargetsPairable(native.payload, old, next);
            return !!a && !!b && sameContainerIdentity({ ...a, type: native.payload.type }, b);
          }
          if (entry.block.kind !== afterBlock.kind) return false;
          if (entry.block.kind === 'table' && afterBlock.kind === 'table') return true;
          if (
            entry.block.kind === 'paragraph' &&
            afterBlock.kind === 'paragraph' &&
            old.type === next.type
          )
            return true;
          if (
            entry.block.kind === 'paragraph' &&
            afterBlock.kind === 'paragraph' &&
            old.type !== next.type
          ) {
            const a = compileInlineTarget(old),
              b = compileInlineTarget(next);
            return (
              a.success &&
              b.success &&
              a.target.plainText !== null &&
              a.target.plainText === b.target.plainText
            );
          }
          return siblingContentPairable(old, next);
        },
      }
    );
    if (!result.success) return result;
    const listTargets = resolveTargetLists(
      { revision, pkg: input.pkg, partName: part.name },
      after,
      result.pairs.map((pair) => ({
        before: mapped[pair.beforeIndex]!.block,
        afterIndex: pair.afterIndex,
        paragraphId: mapped[pair.beforeIndex]!.binding!.structuralOwnerId!,
      }))
    );
    if (!listTargets.success) return listTargets;
    const retained = new Set(result.pairs.map((pair) => pair.afterIndex));
    for (let index = 0; index < after.length; index++) {
      if (retained.has(index)) continue;
      const block = after[index]!,
        ast = astOf(block);
      const problem = insertedBlockProblem(block, ast, listTargets.targets[index]!);
      if (problem) return problem;
    }
    let quoteStyle: string | null | undefined;
    const structural: CanonicalEdit[] = [];
    for (const [index, edit] of result.plan.operations.entries()) {
      const gap = result.gaps[index]!;
      if (edit.kind !== 'siblings' || !edit.inserted.length) {
        structural.push(edit);
        continue;
      }
      const fragments = quoteSiblingInsertion(
        edit.inserted,
        edit.inserted.map((_, insertedIndex) => {
          if (!afterQuotes.quoted.has(after[gap.afterStart + insertedIndex]!)) return null;
          const path = [...afterPath, gap.afterStart + insertedIndex + offset.after];
          return afterQuotes.origins.get(path.join('/'))!.slice(0, -1).join('/');
        })
      );
      const inserted = fragments.inserted;
      const quotes = inserted.flatMap((ast, insertedIndex) =>
        ast.type === 'blockquote' ? [insertedIndex] : []
      );
      if (quotes.length && quoteStyle === undefined)
        quoteStyle = options.quoteStyle ? options.quoteStyle() : documentQuoteStyle(input.pkg);
      if (quotes.length && !quoteStyle)
        return fail('A new block quote needs a Quote paragraph style defined by the document.');
      const nestedBlocks = planInsertedBlocks(
        { revision, pkg: input.pkg, partName: part.name },
        edit.inserted.map((ast, insertedIndex) => ({
          block: after[gap.afterStart + insertedIndex]!,
          ast,
        })),
        options.allocateBlockId,
        options.newBlockPayload
      );
      if (!nestedBlocks.success) return nestedBlocks;
      structural.push({
        ...edit,
        inserted,
        lists: [
          ...edit.inserted.flatMap((ast, insertedIndex) =>
            (ast.type === 'paragraph' || ast.type === 'heading') &&
            !afterQuotes.quoted.has(after[gap.afterStart + insertedIndex]!)
              ? [
                  {
                    path: fragments.path([insertedIndex]),
                    target: listTargets.targets[gap.afterStart + insertedIndex]!,
                  },
                ]
              : []
          ),
          ...nestedBlocks.plan.lists.map((binding) => ({
            ...binding,
            path: fragments.path(binding.path),
          })),
        ],
        ...(nestedBlocks.plan.blocks.length
          ? {
              blocks: nestedBlocks.plan.blocks.map((binding) => ({
                ...binding,
                path: fragments.path(binding.path),
              })),
            }
          : {}),
        ...(quotes.length
          ? { styles: quotes.map((path) => ({ path: [path], styleId: quoteStyle! })) }
          : {}),
      });
    }
    for (const pair of result.pairs) {
      const entry = mapped[pair.beforeIndex]!,
        next = after[pair.afterIndex]!,
        binding: BoundSourceBlock = entry.binding!,
        oldPath = [...beforePath, pair.beforeIndex + offset.before],
        newPath = [...afterPath, pair.afterIndex + offset.after],
        ownerId = binding.structuralOwnerId!;
      pairs.push(
        Object.freeze({
          id: ownerId,
          beforePath: Object.freeze(beforeQuotes.origins.get(oldPath.join('/')) ?? oldPath),
          afterPath: Object.freeze(afterQuotes.origins.get(newPath.join('/')) ?? newPath),
        })
      );
      const wasQuote = beforeQuotes.quoted.has(entry.block),
        isQuote = afterQuotes.quoted.has(next);
      if (deepSame(entry.block, next) && wasQuote === isQuote) continue;
      let membership: readonly CanonicalEdit[] = [];
      if (listKey(entry.block) !== listKey(next)) {
        const list = planListMembership({ revision, pkg: input.pkg, partName: part.name }, [
          { paragraphId: ownerId, target: listTargets.targets[pair.afterIndex]! },
        ]);
        if (!list.success) return list;
        membership = list.plan.operations;
      }
      if (binding.atomicControlId && byControl.has(ownerId)) {
        const block = byControl.get(ownerId)!;
        const atomic = planTargetAtomic(
          { pkg: input.pkg, revision, blockId: block.id },
          block.payload,
          entry.ast,
          astOf(next)
        );
        if (!atomic.success) return atomic;
        operations.push(...atomic.plan.operations);
      } else if (
        entry.block.kind === 'paragraph' &&
        next.kind === 'paragraph' &&
        findNode(part, ownerId)?.kind === 'paragraph'
      ) {
        const paragraph = planTargetParagraph({
          input,
          source,
          binding,
          before: entry.block,
          after: next,
          headingStyle: options.headingStyle,
        });
        if (!paragraph.success) return paragraph;
        operations.push(...paragraph.plan.operations);
        if (wasQuote !== isQuote) {
          const style = isQuote
            ? options.quoteStyle
              ? options.quoteStyle()
              : documentQuoteStyle(input.pkg)
            : null;
          if (isQuote && !style)
            return fail('A block quote needs a Quote paragraph style defined by the document.');
          // Heading transitions already choose their own target style.
          if (isQuote || next.ast.type === 'paragraph') {
            const format = planParagraphFormat(
              { revision, part, paragraphId: ownerId },
              style
                ? { properties: [{ localName: 'pStyle', attributes: { val: style } }] }
                : { properties: [], remove: ['pStyle'] }
            );
            if (!format.success) return format;
            operations.push(...format.plan.operations);
          }
        }
      } else if (
        entry.block.kind === 'table' &&
        next.kind === 'table' &&
        !binding.atomicControlId
      ) {
        const table = planTargetTable({
          input,
          source,
          binding,
          bindings,
          before: entry.block,
          after: next,
          beforePath: oldPath,
          afterPath: newPath,
          astOf,
          targetAst: target.ast,
          same: deepSame,
          visitCell: stream,
        });
        if (!table.success) return table;
        operations.push(...table.plan.operations);
      } else if (
        entry.block.kind === 'container' &&
        next.kind === 'container' &&
        !binding.atomicControlId
      ) {
        const block = byControl.get(ownerId);
        const old = domainNodeFromAst(entry.ast),
          current = domainNodeFromAst(astOf(next));
        if (!block || !old || !current)
          return fail('The retained container needs its application metadata codec.');
        const set: Record<string, JsonObject[string]> = {},
          remove: string[] = [];
        for (const name of new Set([...Object.keys(old), ...Object.keys(current)])) {
          if (name === 'type') continue;
          if (
            Object.hasOwn(old, name) === Object.hasOwn(current, name) &&
            jsonDataKey(old[name] ?? null) === jsonDataKey(current[name] ?? null)
          )
            continue;
          if (Object.hasOwn(current, name))
            Object.defineProperty(set, name, {
              value: current[name],
              enumerable: true,
              configurable: true,
              writable: true,
            });
          else remove.push(name);
        }
        const metadata = planDomainProperties(
          { revision, pkg: input.pkg, blockId: block.id },
          { set, remove }
        );
        if (!metadata.success) return metadata;
        operations.push(...metadata.plan.operations);
        const parents = new Set(
          source.bindings
            .filter(
              (child) =>
                child.logicalPath.length === oldPath.length + 1 &&
                oldPath.every((value, index) => child.logicalPath[index] === value)
            )
            .map(
              (child) =>
                child.structuralOwnerId &&
                source.index.owners.find((owner) => owner.id === child.structuralOwnerId)?.parentId
            )
            .filter((id): id is string => !!id)
        );
        if (entry.block.children.length !== 0 && parents.size !== 1)
          return fail('Retained container content needs resolved native child windows.');
        const childParent = [...parents][0];
        if (!childParent) {
          if (next.children.length)
            return fail('Empty container insertion needs native seed ownership.');
        } else {
          const error = stream(entry.block.children, next.children, oldPath, newPath, childParent);
          if (error) return error;
        }
      } else return fail('This retained target needs its table or atomic payload dispatcher.');
      operations.push(...membership);
    }
    operations.push(...structural);
    return null;
  };
  const error = stream(source.logical.blocks, target.blocks, [], [], source.index.bodyId);
  if (error) return error;
  const result = createEditPlan(revision, operations);
  return result.success ? { ...result, pairs: Object.freeze(pairs) } : result;
}
