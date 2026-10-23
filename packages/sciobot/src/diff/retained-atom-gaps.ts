import { targetSourceRangeAt, type TargetAstNode } from '../projection/parse-target';
import { paragraphTextOf, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { BoundLogicalSource, BoundSourceBlock } from './bound-source-types';
import type { CompiledInlineTarget, InlineTargetToken } from './inline-target';
import { targetAstKey } from './target-key';

function atomKey(ast: TargetAstNode): string {
  if (ast.name === 'source' && Array.isArray(ast.attributes)) {
    const attributes = ast.attributes as { type?: unknown; name?: unknown; value?: unknown }[];
    if (
      attributes.every(
        (attr) =>
          attr.type === 'mdxJsxAttribute' &&
          typeof attr.name === 'string' &&
          (typeof attr.value === 'string' || attr.value === null)
      ) &&
      new Set(attributes.map((attr) => attr.name)).size === attributes.length
    )
      return targetAstKey({
        ...ast,
        attributes: [...attributes].sort((a, b) => String(a.name).localeCompare(String(b.name))),
      } as TargetAstNode);
  }
  return targetAstKey(ast);
}
type Atom = Extract<InlineTargetToken, { kind: 'atom' }>;
export interface RetainedAtomGap {
  readonly before: TargetAstNode;
  readonly after: TargetAstNode;
  readonly start: number;
  readonly end: number;
}
function sliceAst(ast: TargetAstNode, tokens: readonly InlineTargetToken[]): TargetAstNode {
  interface Paths {
    leaf: boolean;
    children: Map<number, Paths>;
  }
  const paths: Paths = { leaf: false, children: new Map() };
  for (const token of tokens) {
    let cursor = paths;
    for (const index of token.path) {
      let next = cursor.children.get(index);
      if (!next) {
        next = { leaf: false, children: new Map() };
        cursor.children.set(index, next);
      }
      cursor = next;
    }
    cursor.leaf = true;
  }
  const visit = (node: TargetAstNode, selected: Paths): TargetAstNode => {
    if (selected.leaf) return node;
    const original = (node.children ?? []) as TargetAstNode[];
    const children = [...selected.children].map(([index, child]) => visit(original[index]!, child));
    return { ...node, children };
  };
  const { depth: _depth, ...rest } = visit(ast, paths);
  return { ...rest, type: 'paragraph' };
}
/** Existing controls remain native anchors. Only the text gaps between them are editable here. */
export function retainedAtomGaps(
  part: OoxmlPart,
  paragraphId: string,
  source: BoundLogicalSource,
  binding: BoundSourceBlock,
  before: CompiledInlineTarget,
  after: CompiledInlineTarget,
  base: number
): readonly RetainedAtomGap[] | null {
  if (before.hyperlinks.length || after.hyperlinks.length) return null;
  const tokenKey = (atom: Atom) => JSON.stringify([atomKey(atom.ast), atom.marks]);
  const oldAtoms = before.tokens.flatMap((token, index) =>
    token.kind === 'atom' ? [{ atom: token, index }] : []
  );
  if (!oldAtoms.length) return null;
  const counts = new Map<string, number>();
  for (const { atom } of oldAtoms)
    counts.set(tokenKey(atom), (counts.get(tokenKey(atom)) ?? 0) + 1);
  const targets = new Map<string, { readonly atom: Atom; readonly index: number }[]>();
  after.tokens.forEach((token, index) => {
    if (token.kind !== 'atom') return;
    const key = tokenKey(token),
      entries = targets.get(key) ?? [];
    entries.push({ atom: token, index });
    targets.set(key, entries);
  });
  const ranges = new Map<string, (typeof binding.inline)[number][]>();
  for (const inline of binding.inline) {
    const key = `${inline.start}:${inline.end}`,
      entries = ranges.get(key) ?? [];
    entries.push(inline);
    ranges.set(key, entries);
  }
  const matches: { old: number; next: number; start: number; end: number }[] = [];
  const cursors = new Map<string, number>();
  let nextAt = 0;
  for (const { atom, index } of oldAtoms) {
    const key = tokenKey(atom),
      candidates = targets.get(key) ?? [];
    let cursor = cursors.get(key) ?? 0;
    while (cursor < candidates.length && candidates[cursor]!.index < nextAt) cursor++;
    // Equal repeated atoms pair by ordinal only when their counts agree.
    if (
      cursor === candidates.length ||
      (candidates.length - cursor > 1 && counts.get(key) !== candidates.length)
    )
      return null;
    const next = candidates[cursor]!;
    cursors.set(key, cursor + 1);
    const position = atom.ast.position as
      | { start?: { offset?: number }; end?: { offset?: number } }
      | undefined;
    if (position?.start?.offset === undefined || position.end?.offset === undefined) return null;
    const range = targetSourceRangeAt(source.parsed, position.start.offset, position.end.offset);
    const native = range && ranges.get(`${range.start}:${range.end}`);
    if (!native || native.length !== 1) return null;
    matches.push({
      old: index,
      next: next.index,
      start: native[0]!.native.start,
      end: native[0]!.native.end,
    });
    nextAt = next.index + 1;
  }
  const nativeText = paragraphTextOf(part, paragraphId);
  if (nativeText === null) return null;
  const gaps: RetainedAtomGap[] = [];
  let oldAt = 0,
    targetAt = 0,
    start = base;
  for (const match of [
    ...matches,
    {
      old: before.tokens.length,
      next: after.tokens.length,
      start: nativeText.length,
      end: nativeText.length,
    },
  ]) {
    if (match.start < start || match.end < match.start) return null;
    const old = before.tokens.slice(oldAt, match.old),
      target = after.tokens.slice(targetAt, match.next);
    if (
      old.some((token) => token.kind !== 'text') ||
      old.map((token) => (token.kind === 'text' ? token.value : '')).join('') !==
        nativeText.slice(start, match.start)
    )
      return null;
    gaps.push({
      before: sliceAst(before.ast, old),
      after: sliceAst(after.ast, target),
      start,
      end: match.start,
    });
    oldAt = match.old + 1;
    targetAt = match.next + 1;
    start = match.end;
  }
  return gaps;
}
