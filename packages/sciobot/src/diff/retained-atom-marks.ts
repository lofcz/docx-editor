import { retainedAtomKey } from './retained-atom-gaps';
import { paragraphTextOf, type OoxmlPart } from '@docx-editor.dev/core/store';
import { targetSourceRangeAt } from '../projection/parse-target';
import { hyperlinkCharacterStyle, inlineCodeCharacterStyles } from '../native/inline-code';
import type { TargetParagraphSource } from './target-inline';
import type { CompiledInlineTarget } from './inline-target';
import { planMarkDeltas } from './insertion-marks';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';
import { jsonDataKey } from './target-key';

/** Format unchanged atoms through their exact native spans, without replacing their controls. */
export function planRetainedAtomMarks(
  context: TargetParagraphSource,
  part: OoxmlPart,
  paragraphId: string,
  before: CompiledInlineTarget,
  after: CompiledInlineTarget,
  base: number
): EditPlanResult | null {
  if (before.tokens.length !== after.tokens.length || !before.tokens.some((t) => t.kind === 'atom'))
    return null;
  if (
    jsonDataKey(before.hyperlinks.map(({ path: _path, ...link }) => link)) !==
    jsonDataKey(after.hyperlinks.map(({ path: _path, ...link }) => link))
  )
    return null;
  const text = paragraphTextOf(part, paragraphId);
  if (text === null) return null;
  const operations: CanonicalEdit[] = [];
  let offset = base;
  for (let i = 0; i < before.tokens.length; i++) {
    const old = before.tokens[i]!,
      next = after.tokens[i]!;
    if (old.kind !== next.kind) return null;
    const linkOwner = (target: CompiledInlineTarget, path: readonly number[] | null) =>
      path === null
        ? -1
        : target.hyperlinks.findIndex((link) => link.path.join('/') === path.join('/'));
    if (linkOwner(before, old.hyperlink) !== linkOwner(after, next.hyperlink)) return null;
    let end: number;
    if (old.kind === 'text' && next.kind === 'text') {
      if (old.value !== next.value || text.slice(offset, offset + old.value.length) !== old.value)
        return null;
      end = offset + old.value.length;
    } else if (old.kind === 'atom' && next.kind === 'atom') {
      if (retainedAtomKey(old.ast) !== retainedAtomKey(next.ast)) return null;
      const position = old.ast.position as
        | { start?: { offset?: number }; end?: { offset?: number } }
        | undefined;
      if (position?.start?.offset === undefined || position.end?.offset === undefined) return null;
      const range = targetSourceRangeAt(
        context.source.parsed,
        position.start.offset,
        position.end.offset
      );
      const native =
        range &&
        context.binding.inline.filter(
          (item) => item.start === range.start && item.end === range.end
        );
      if (!native || native.length !== 1 || native[0]!.native.start !== offset) return null;
      end = native[0]!.native.end;
    } else return null;
    const value = text.slice(offset, end);
    if (old.marks.join('/') !== next.marks.join('/')) {
      if (!value.length) return null;
      const planned = planMarkDeltas(
        {
          revision: context.source.index.revision,
          part,
          paragraphId,
          base: offset,
          length: value.length,
          displayMode: context.input.projection.displayMode!,
          hyperlinkStyle: hyperlinkCharacterStyle(context.input.pkg),
          codeStyles: inlineCodeCharacterStyles(context.input.pkg),
        },
        [{ value, marks: old.marks }],
        [{ value, marks: next.marks }]
      );
      if (!planned.success) return planned;
      operations.push(...planned.operations);
    }
    offset = end;
  }
  if (offset !== (context.nativeRange?.end ?? text.length)) return null;
  return createEditPlan(context.source.index.revision, operations);
}
