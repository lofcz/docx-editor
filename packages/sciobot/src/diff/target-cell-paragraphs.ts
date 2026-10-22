import { diffChars } from 'diff';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { dataFingerprint } from '../projection/protected-snapshots';
import { targetSourceRangeAt, type TargetAstNode } from '../projection/parse-target';
import type { BoundLogicalSource, BoundSourceBlock } from './bound-source-types';
import type { LogicalBlock, LogicalBlockDocument, LogicalParagraph } from './logical-block-types';
import type { TargetSourceInput } from './target-source-types';
import { compileInlineTarget } from './inline-target';
import { targetAstKey } from './target-key';
import { alignTargetSequence } from './target-sequence';
import { SIBLING_BLOCK_LIMITS } from './sibling-blocks';

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
const children = (node: TargetAstNode) => (node.children ?? []) as readonly TargetAstNode[];
const isBreak = (node: TargetAstNode) =>
  node.type === 'break' ||
  (node.type === 'mdxJsxTextElement' &&
    node.name === 'br' &&
    Array.isArray(node.attributes) &&
    !node.attributes.length &&
    !children(node).length);
function at(root: TargetAstNode, path: readonly number[]): TargetAstNode {
  let node = root;
  for (const index of path) node = children(node)[index]!;
  return node;
}
function inlineBreaks(paragraph: LogicalParagraph) {
  const compiled = compileInlineTarget(paragraph.ast);
  if (!compiled.success || compiled.target.plainText === null) return null;
  let offset = 0;
  const breaks: { path: readonly number[]; node: TargetAstNode; offset: number }[] = [];
  for (const token of compiled.target.tokens) {
    if (token.kind !== 'text') return null;
    const node = at(paragraph.ast, token.path);
    if (isBreak(node)) breaks.push({ path: token.path, node, offset });
    offset += token.value.length;
  }
  return { text: compiled.target.plainText, breaks };
}

/** Split only selected syntax nodes, retaining inline wrapper metadata and original AST paths. */
function split(paragraph: LogicalParagraph, separators: ReadonlySet<string>): LogicalParagraph[] {
  const walk = (node: TargetAstNode, path: readonly number[]): TargetAstNode[][] => {
    if (separators.has(path.join('/'))) return [[], []];
    if (!children(node).length) return [[node]];
    const groups: TargetAstNode[][] = [[]];
    children(node).forEach((child, index) => {
      const parts = walk(child, [...path, index]);
      groups.at(-1)!.push(...parts[0]!);
      for (const part of parts.slice(1)) groups.push([...part]);
    });
    return groups.map((group) => (group.length ? [{ ...node, children: group }] : []));
  };
  const groups: TargetAstNode[][] = [[]];
  children(paragraph.ast).forEach((child, index) => {
    const parts = walk(child, [index]);
    groups.at(-1)!.push(...parts[0]!);
    for (const part of parts.slice(1)) groups.push([...part]);
  });
  return groups.map((nodes) => {
    const ast: TargetAstNode = { type: 'paragraph', children: nodes };
    return {
      ...paragraph,
      ast,
      key: `paragraph:${dataFingerprint({ body: targetAstKey(ast), html: null })}`,
    };
  });
}

/** Expand source-proven GFM paragraph separators before ordinary native sibling planning. */
export function prepareTargetCellParagraphs(
  input: TargetSourceInput,
  source: BoundLogicalSource,
  target: LogicalBlockDocument
):
  | {
      readonly success: true;
      readonly source: BoundLogicalSource;
      readonly target: LogicalBlockDocument;
    }
  | NativeEditFailure {
  if (targetAstKey(source.logical.ast) === targetAstKey(target.ast))
    return { success: true, source, target };
  if (!source.bindings.some((binding) => binding.table?.cellId && binding.contributions.length > 1))
    return { success: true, source, target };
  const original = new Map(
    source.bindings.map((binding) => [binding.logicalPath.join('/'), binding])
  );
  const replacements = new Map<string, readonly BoundSourceBlock[]>();
  const owners = new Map(source.index.owners.map((owner) => [owner.id, owner]));
  const part = input.pkg.parts.get(source.index.partName)!;
  let error: NativeEditFailure | null = null;

  const cell = (before: LogicalParagraph, after: LogicalParagraph, path: readonly number[]) => {
    const binding = original.get(path.join('/'));
    if (!binding?.table?.cellId || binding.contributions.length < 2) return null;
    if (
      !before.synthetic ||
      before.list ||
      before.prefixRemoved ||
      before.html ||
      after.list ||
      after.html
    )
      return null;
    const contributions = binding.contributions;
    if (new Set(contributions.map((item) => item.ownerId)).size !== contributions.length)
      return null;
    const old = inlineBreaks(before),
      next = inlineBreaks(after);
    if (!old || !next) return null;
    const generated = new Set<string>();
    for (let index = 0; index < contributions.length - 1; index++) {
      const left = contributions[index]!,
        right = contributions[index + 1]!;
      const candidates = old.breaks.filter((item) => {
        const position = item.node.position as
          | { start?: { offset?: number }; end?: { offset?: number } }
          | undefined;
        if (position?.start?.offset === undefined || position.end?.offset === undefined)
          return false;
        const range = targetSourceRangeAt(
          source.parsed,
          position.start.offset,
          position.end.offset
        );
        return range && left.end <= range.start && range.end <= right.start;
      });
      if (candidates.length !== 1) return null;
      generated.add(candidates[0]!.path.join('/'));
    }
    const oldParts = split(before, generated);
    if (oldParts.length !== contributions.length) return null;
    for (const [index, paragraph] of oldParts.entries()) {
      const contribution = contributions[index]!,
        owner = owners.get(contribution.ownerId);
      const compiled = compileInlineTarget(paragraph.ast);
      if (
        !owner ||
        owner.kind !== 'paragraph' ||
        owner.parentId !== binding.table.cellId ||
        owner.extents.length !== 1 ||
        !compiled.success ||
        compiled.target.plainText !== paragraphTextOf(part, owner.id)
      )
        return null;
    }
    // Retain authored breaks only where the text diff carries their original newline forward.
    // A new separator in a multi-paragraph GFM cell introduces a new native paragraph.
    const manual = new Set(
      old.breaks.filter((item) => !generated.has(item.path.join('/'))).map((item) => item.offset)
    );
    const retainedManual = new Set<number>();
    const changes = diffChars(old.text, next.text, { timeout: SIBLING_BLOCK_LIMITS.diffTimeoutMs });
    if (!changes) {
      error = fail('Cell separator matching exceeded its planning budget.');
      return null;
    }
    let oldOffset = 0,
      newOffset = 0;
    for (const change of changes) {
      if (!change.added && !change.removed) {
        for (const offset of manual)
          if (oldOffset <= offset && offset < oldOffset + change.value.length)
            retainedManual.add(newOffset + offset - oldOffset);
      }
      if (!change.added) oldOffset += change.value.length;
      if (!change.removed) newOffset += change.value.length;
    }
    const newParts = split(
      after,
      new Set(
        next.breaks
          .filter((item) => !retainedManual.has(item.offset))
          .map((item) => item.path.join('/'))
      )
    );
    const prefix = path.slice(0, -1);
    replacements.set(
      path.join('/'),
      oldParts.map((paragraph, index) => {
        const contribution = contributions[index]!;
        return {
          ...binding,
          logicalPath: [...prefix, index],
          astPath: paragraph.path,
          source: {
            start: contribution.start,
            end: contribution.end,
            exact: true,
            generated: false,
          },
          contributions: [contribution],
          structuralOwnerId: contribution.ownerId,
          inline: binding.inline.filter(
            (item) => contribution.start <= item.start && item.end <= contribution.end
          ),
        };
      })
    );
    return { before: oldParts, after: newParts };
  };
  const streams = (
    before: readonly LogicalBlock[],
    after: readonly LogicalBlock[],
    path: readonly number[]
  ): { before: readonly LogicalBlock[]; after: readonly LogicalBlock[] } => {
    const aligned = alignTargetSequence(
      before.map((block) => block.key),
      after.map((block) => block.key)
    );
    if (!aligned.success) {
      error = aligned;
      return { before, after };
    }
    const old = [...before],
      next = [...after];
    for (const pair of aligned.alignment.pairs) {
      const a = before[pair.before]!,
        b = after[pair.after]!;
      if (!('children' in a) || !('children' in b) || a.kind !== b.kind) continue;
      const childPath = [...path, pair.before];
      if (
        a.kind === 'cell' &&
        a.header.type === 'tableCell' &&
        b.header.type === 'tableCell' &&
        a.children.length === 1 &&
        b.children.length === 1 &&
        a.children[0]!.kind === 'paragraph' &&
        b.children[0]!.kind === 'paragraph'
      ) {
        const expanded = cell(a.children[0], b.children[0], [...childPath, 0]);
        if (expanded) {
          old[pair.before] = { ...a, children: expanded.before };
          next[pair.after] = { ...b, children: expanded.after };
        }
      } else {
        const expanded = streams(a.children, b.children, childPath);
        old[pair.before] = { ...a, children: expanded.before };
        next[pair.after] = { ...b, children: expanded.after };
      }
    }
    return { before: old, after: next };
  };
  const expanded = streams(source.logical.blocks, target.blocks, []);
  if (error) return error;
  if (!replacements.size) return { success: true, source, target };
  const bindings = source.bindings.flatMap(
    (binding) => replacements.get(binding.logicalPath.join('/')) ?? [binding]
  );
  const windows = source.windows.map((window) => ({
    ...window,
    members: window.members.map((member) => ({
      ...member,
      logicalPaths: member.logicalPaths.flatMap((path) => {
        const replaced = replacements.get(path.join('/'));
        return replaced
          ? replaced
              .filter((binding) => binding.structuralOwnerId === member.ownerId)
              .map((binding) => binding.logicalPath)
          : [path];
      }),
    })),
  }));
  return {
    success: true,
    source: {
      ...source,
      logical: { ...source.logical, blocks: expanded.before },
      bindings,
      windows,
    },
    target: { ...target, blocks: expanded.after },
  };
}
