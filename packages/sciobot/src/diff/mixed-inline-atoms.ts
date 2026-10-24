import { alignRetainedAtoms } from './retained-atom-alignment';
import { paragraphTextOf, type OoxmlPart } from '@docx-editor.dev/core/store';
import { targetSourceRangeAt, type TargetAstNode } from '../projection/parse-target';
import type { TargetParagraphSource } from './target-inline';
import type { CompiledInlineTarget } from './inline-target';
import { retainedAtomKey } from './retained-atom-gaps';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';
import { planParagraphText } from './paragraph-text';

/** Treat retained atoms as native anchors, never as text that a lexical diff can replace. */
export function mixedInlineAtoms(
  context: TargetParagraphSource,
  part: OoxmlPart,
  paragraphId: string,
  before: CompiledInlineTarget,
  after: CompiledInlineTarget,
  base: number
): {
  before: TargetAstNode;
  after: TargetAstNode;
  text: EditPlanResult;
  anchors: readonly { start: number; end: number }[];
} | null {
  if (context.nativeRange) return null;
  const oldAtoms = before.tokens.filter((token) => token.kind === 'atom');
  const nextAtoms = after.tokens.filter((token) => token.kind === 'atom');
  if (!oldAtoms.length || nextAtoms.length < oldAtoms.length) return null;
  const text = paragraphTextOf(part, paragraphId);
  if (text === null) return null;
  const keys = nextAtoms.map((token) => retainedAtomKey(token.ast));
  const oldKeys = oldAtoms.map((token) => retainedAtomKey(token.ast));
  const matches: { old: string; next: string; start: number; end: number; value: string }[] = [];
  const counts = (values: readonly string[]) => {
    const result = new Map<string, number>();
    for (const value of values) result.set(value, (result.get(value) ?? 0) + 1);
    return result;
  };
  const oldCounts = counts(oldKeys),
    nextCounts = counts(keys);
  if ([...oldCounts].some(([key, count]) => (nextCounts.get(key) ?? 0) < count)) return null;
  const repeated = [...oldCounts].some(([key, count]) => nextCounts.get(key) !== count);
  const alignment =
    context.retainedAtomAlignment ?? (repeated ? alignRetainedAtoms(before, after) : null);
  if ((repeated && !alignment) || (alignment && alignment.length !== oldAtoms.length)) return null;
  const candidates = new Map<string, number[]>();
  keys.forEach((key, index) => {
    const entries = candidates.get(key) ?? [];
    entries.push(index);
    candidates.set(key, entries);
  });
  const ordinals = new Map<string, number>();
  const ranges = new Map<string, (typeof context.binding.inline)[number][]>();
  for (const inline of context.binding.inline) {
    const key = `${inline.start}:${inline.end}`,
      entries = ranges.get(key) ?? [];
    entries.push(inline);
    ranges.set(key, entries);
  }
  let nextAt = 0;
  for (let i = 0; i < oldAtoms.length; i++) {
    const atom = oldAtoms[i]!,
      key = oldKeys[i]!;
    const ordinal = ordinals.get(key) ?? 0;
    const next = alignment?.[i] ?? candidates.get(key)?.[ordinal];
    if (next === undefined || next < nextAt || keys[next] !== key) return null;
    ordinals.set(key, ordinal + 1);
    nextAt = next + 1;
    const position = atom.ast.position as
      | { start?: { offset?: number }; end?: { offset?: number } }
      | undefined;
    if (position?.start?.offset === undefined || position.end?.offset === undefined) return null;
    const range = targetSourceRangeAt(
      context.source.parsed,
      position.start.offset,
      position.end.offset
    );
    const native = range && ranges.get(`${range.start}:${range.end}`);
    if (!native || native.length !== 1) return null;
    const { start, end } = native[0]!.native;
    if (start < (matches.at(-1)?.end ?? base) || end <= start || end > text.length) return null;
    matches.push({
      old: atom.path.join('/'),
      next: nextAtoms[next]!.path.join('/'),
      start,
      end,
      value: text.slice(start, end),
    });
  }
  const replace = (ast: TargetAstNode, side: 'old' | 'next'): TargetAstNode => {
    const replacements = new Map(matches.map((match) => [match[side], match.value]));
    const visit = (node: TargetAstNode, path: readonly number[]): TargetAstNode => {
      const value = replacements.get(path.join('/'));
      if (value !== undefined) return { type: 'text', value };
      if (!node.children) return node;
      return {
        ...node,
        children: (node.children as TargetAstNode[]).map((child, i) => visit(child, [...path, i])),
      };
    };
    return visit(ast, []);
  };
  const beforePositions = new Map(
    before.tokens.map((token, index) => [token.path.join('/'), index])
  );
  const afterPositions = new Map(after.tokens.map((token, index) => [token.path.join('/'), index]));
  const operations: CanonicalEdit[] = [];
  let oldAt = 0,
    afterAt = 0,
    start = base;
  for (const match of [
    ...matches,
    { old: '', next: '', start: text.length, end: text.length, value: '' },
  ]) {
    const oldEnd = match.old ? beforePositions.get(match.old)! : before.tokens.length;
    const afterEnd = match.next ? afterPositions.get(match.next)! : after.tokens.length;
    const oldGap = before.tokens.slice(oldAt, oldEnd);
    const newGap = after.tokens.slice(afterAt, afterEnd);
    if (oldGap.some((token) => token.kind !== 'text')) return null;
    const oldText = oldGap.map((token) => (token.kind === 'text' ? token.value : '')).join('');
    if (oldText !== text.slice(start, match.start)) return null;
    const planned = planParagraphText(
      {
        revision: context.source.index.revision,
        partName: part.name,
        paragraphId,
        offset: start,
        text: oldText,
      },
      newGap.map((token) => (token.kind === 'text' ? token.value : '')).join('')
    );
    if (!planned.success)
      return {
        before: replace(before.ast, 'old'),
        after: replace(after.ast, 'next'),
        text: planned,
        anchors: matches,
      };
    operations.unshift(...planned.plan.operations);
    oldAt = oldEnd + 1;
    afterAt = afterEnd + 1;
    start = match.end;
  }
  return {
    before: replace(before.ast, 'old'),
    after: replace(after.ast, 'next'),
    text: createEditPlan(context.source.index.revision, operations),
    anchors: matches,
  };
}
