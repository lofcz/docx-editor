import {
  canonicalOoxmlFingerprint,
  findNode,
  withPart,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { targetSourceRangeAt, type TargetAstNode } from '../projection/parse-target';
import { lowerInlineRemoval } from '../native/remove-inline-atom';
import { retainedAtomKey } from './retained-atom-gaps';
import { alignRetainedAtoms } from './retained-atom-alignment';
import type { CompiledInlineTarget } from './inline-target';
import type { TargetParagraphSource } from './target-inline';
import type { CanonicalEdit, EditPlanResult } from './plan';

/** Remove addressed atoms first, then plan all other changes in the resulting native coordinates. */
export function planInlineRemovals(
  context: TargetParagraphSource,
  part: OoxmlPart,
  paragraphId: string,
  before: CompiledInlineTarget,
  after: CompiledInlineTarget
):
  | { readonly context: TargetParagraphSource; readonly edits: readonly CanonicalEdit[] }
  | EditPlanResult
  | null {
  if (context.nativeRange || context.retainedAtomAlignment || before.plainText !== null)
    return null;
  const oldAtoms = before.tokens.filter((token) => token.kind === 'atom');
  const key = (ast: TargetAstNode) =>
    ast.type === 'image' ? 'image' : ast.name === 'source' ? 'source' : retainedAtomKey(ast);
  const remaining = new Map<string, number>();
  for (const token of after.tokens)
    if (token.kind === 'atom')
      remaining.set(key(token.ast), (remaining.get(key(token.ast)) ?? 0) + 1);
  let missing = false;
  for (const token of oldAtoms) {
    const count = remaining.get(key(token.ast)) ?? 0;
    if (!count) missing = true;
    else remaining.set(key(token.ast), count - 1);
  }
  if (!missing) return null;
  const alignment = alignRetainedAtoms(before, after, true, true, true);
  if (!alignment || !alignment.some((index) => index < 0)) return null;
  const removedPaths = new Set<string>(),
    removedIds = new Set<string>();
  const spans: { start: number; end: number }[] = [],
    edits: CanonicalEdit[] = [];
  const blocks = buildBlockIndex(context.input.pkg);
  let preview = part;
  try {
    for (let i = 0; i < oldAtoms.length; i++) {
      if (alignment[i]! >= 0) continue;
      const atom = oldAtoms[i]!,
        position = atom.ast.position as
          | { start?: { offset?: number }; end?: { offset?: number } }
          | undefined;
      if (position?.start?.offset === undefined || position.end?.offset === undefined)
        throw Error('The removed atom has no bound source range.');
      const range = targetSourceRangeAt(
        context.source.parsed,
        position.start.offset,
        position.end.offset
      );
      const bindings =
        range &&
        context.binding.inline.filter(
          (entry) => entry.start === range.start && entry.end === range.end
        );
      if (!bindings || bindings.length !== 1)
        throw Error('The removed atom has no unique native control.');
      const binding = bindings[0]!,
        block = blocks.byId.get(binding.blockId),
        control = block && findNode(preview, block.controlId);
      if (!block || !control) throw Error('The removed atom has no native payload.');
      const edit: CanonicalEdit = {
        kind: 'inline-remove',
        partName: part.name,
        paragraphId,
        blockId: block.id,
        controlId: block.controlId,
        before: block.payload,
        fingerprint: canonicalOoxmlFingerprint(control),
      };
      const lowered = lowerInlineRemoval(withPart(context.input.pkg, preview), preview, edit);
      preview = lowered.part;
      edits.push(edit);
      spans.push(binding.native);
      removedPaths.add(atom.path.join('/'));
      removedIds.add(block.id);
    }
    const visit = (node: TargetAstNode, path: number[]): TargetAstNode | null => {
      if (removedPaths.has(path.join('/'))) return null;
      if (!Array.isArray(node.children) || node.children.length === 0) return node;
      const children = (node.children as TargetAstNode[]).flatMap((child, i) => {
        const next = visit(child, [...path, i]);
        return next ? [next] : [];
      });
      if (path.length && !children.length) return null;
      return { ...node, children };
    };
    const shift = (offset: number) =>
      offset -
      spans.reduce((sum, span) => sum + (span.end <= offset ? span.end - span.start : 0), 0);
    return {
      edits,
      context: {
        ...context,
        input: { ...context.input, pkg: withPart(context.input.pkg, preview) },
        before: { ...context.before, ast: visit(before.ast, [])! },
        binding: {
          ...context.binding,
          inline: context.binding.inline
            .filter((entry) => !removedIds.has(entry.blockId))
            .map((entry) => ({
              ...entry,
              native: { start: shift(entry.native.start), end: shift(entry.native.end) },
            })),
        },
        retainedAtomAlignment: alignment.filter((index) => index >= 0),
      },
    };
  } catch (error) {
    return {
      success: false,
      code: 'unsupported-operation',
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
