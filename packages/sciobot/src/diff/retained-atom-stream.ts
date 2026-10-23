import { paragraphTextOf, type OoxmlPart } from '@docx-editor.dev/core/store';
import { targetSourceRangeAt } from '../projection/parse-target';
import type { TargetParagraphSource } from './target-inline';
import type { CompiledInlineTarget, InlineTargetToken } from './inline-target';
import { retainedAtomKey } from './retained-atom-gaps';

/** Resolve native widths without flattening atom identity or their fallback into editable text. */
export function retainedAtomStreams(
  context: TargetParagraphSource,
  part: OoxmlPart,
  paragraphId: string,
  before: CompiledInlineTarget,
  after: CompiledInlineTarget,
  base: number
): { before: CompiledInlineTarget; after: CompiledInlineTarget } | null {
  const text = paragraphTextOf(part, paragraphId);
  if (text === null || context.nativeRange) return null;
  const atoms: { key: string; value: string }[] = [];
  const oldTokens: InlineTargetToken[] = [];
  let offset = base;
  for (const token of before.tokens) {
    if (token.kind === 'text') {
      if (text.slice(offset, offset + token.value.length) !== token.value) return null;
      oldTokens.push(token);
      offset += token.value.length;
      continue;
    }
    const position = token.ast.position as
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
      context.binding.inline.filter((item) => item.start === range.start && item.end === range.end);
    if (!native || native.length !== 1 || native[0]!.native.start !== offset) return null;
    const value = text.slice(offset, native[0]!.native.end);
    atoms.push({ key: retainedAtomKey(token.ast), value });
    oldTokens.push({ ...token, kind: 'text', value });
    offset = native[0]!.native.end;
  }
  if (offset !== text.length || !atoms.length) return null;
  const newTokens: InlineTargetToken[] = [];
  let atomAt = 0;
  for (const token of after.tokens) {
    if (token.kind === 'text') {
      newTokens.push(token);
      continue;
    }
    const atom = atoms[atomAt++];
    if (!atom || atom.key !== retainedAtomKey(token.ast)) return null;
    newTokens.push({ ...token, kind: 'text', value: atom.value });
  }
  const plain = (tokens: readonly InlineTargetToken[]) =>
    tokens.map((t) => (t.kind === 'text' ? t.value : '')).join('');
  if (atomAt !== atoms.length || plain(newTokens) !== text.slice(base)) return null;
  return {
    before: { ...before, tokens: oldTokens, plainText: text.slice(base) },
    after: { ...after, tokens: newTokens, plainText: text.slice(base) },
  };
}
