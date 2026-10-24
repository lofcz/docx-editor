import { alignRetainedAtoms } from './retained-atom-alignment';
import { findNode, type OoxmlNode, type OoxmlPart } from '@docx-editor.dev/core/store';
import { decodeInlineImagePayload } from '../blocks/inline-image-codec';
import { encodeBlockPayload } from '../blocks/payload-schema';
import { buildBlockIndex } from '../blocks/block-index';
import { targetSourceRangeAt, type TargetAstNode } from '../projection/parse-target';
import type { TargetParagraphSource } from './target-inline';
import type { CompiledInlineTarget } from './inline-target';
import { compileInsertedAtoms } from './inserted-atoms';
import { targetAstKey } from './target-key';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';

/** Keep changed image controls native, then let the ordinary planners handle their text, marks and links. */
export function planImageReplacements(
  context: TargetParagraphSource,
  part: OoxmlPart,
  paragraphId: string,
  before: CompiledInlineTarget,
  after: CompiledInlineTarget
):
  | { ast: TargetAstNode; edits: readonly CanonicalEdit[]; alignment: readonly number[] }
  | EditPlanResult
  | null {
  const oldAtoms = before.tokens.filter((token) => token.kind === 'atom');
  const newAtoms = after.tokens.filter((token) => token.kind === 'atom');
  if (oldAtoms.length > newAtoms.length || !oldAtoms.some((atom) => atom.ast.type === 'image'))
    return null;
  const alignment =
    context.retainedAtomAlignment ??
    (oldAtoms.length === newAtoms.length
      ? oldAtoms.map((_, index) => index)
      : alignRetainedAtoms(before, after, true));
  if (
    !alignment ||
    alignment.length !== oldAtoms.length ||
    alignment.some(
      (index, ordinal) =>
        !Number.isInteger(index) ||
        index < 0 ||
        index >= newAtoms.length ||
        (ordinal > 0 && index <= alignment[ordinal - 1]!)
    )
  )
    return null;
  const changes = oldAtoms.flatMap((old, index) => {
    const next = newAtoms[alignment[index]!]!;
    return old.ast.type === 'image' &&
      next.ast.type === 'image' &&
      targetAstKey(old.ast) !== targetAstKey(next.ast)
      ? [{ old, next }]
      : [];
  });
  if (!changes.length) return null;
  const blocks = buildBlockIndex(context.input.pkg),
    edits: CanonicalEdit[] = [];
  const replacements = new Map<string, TargetAstNode>();
  const fail = (error: string): EditPlanResult => ({
    success: false,
    code: 'unsupported-operation',
    error,
  });
  for (const { old, next } of changes) {
    const position = old.ast.position as
      | { start?: { offset?: number }; end?: { offset?: number } }
      | undefined;
    if (position?.start?.offset === undefined || position.end?.offset === undefined)
      return fail('The replaced image has no bound source range.');
    const range = targetSourceRangeAt(
      context.source.parsed,
      position.start.offset,
      position.end.offset
    );
    const native =
      range &&
      context.binding.inline.filter(
        (entry) => entry.start === range.start && entry.end === range.end
      );
    if (!native || native.length !== 1)
      return fail('The replaced image has no unique native control.');
    const block = blocks.byId.get(native[0]!.blockId);
    const payload = block && decodeInlineImagePayload(encodeBlockPayload(block.payload));
    const control = block && findNode(part, block.controlId);
    if (!block || !payload || control?.kind !== 'contentControl')
      return fail('The replaced image is not a managed native picture.');
    if (
      payload.attributes.reference !== old.ast.url ||
      payload.attributes.alt !== (old.ast.alt ?? '') ||
      payload.attributes.title !== (old.ast.title ?? null)
    )
      return fail('The image source no longer matches its native payload.');
    const drawings: OoxmlNode[] = [];
    const visit = (node: OoxmlNode) => {
      if (node.kind === 'drawing') drawings.push(node);
      else if (node.kind !== 'textValue') node.children.forEach(visit);
    };
    visit(control);
    if (drawings.length !== 1) return fail('The managed image needs exactly one native drawing.');
    const compiled = compileInsertedAtoms(
      { location: { kind: 'retained-paragraph', paragraphId }, offset: 0 },
      { type: 'paragraph', children: [next.ast] }
    );
    if (!compiled.success) return compiled;
    const atom = compiled.plan.operations[0];
    if (compiled.plan.operations.length !== 1 || atom?.kind !== 'drawing')
      return fail('The target image has no explicit resource intention.');
    edits.push({
      kind: 'inline-image',
      partName: part.name,
      paragraphId,
      blockId: block.id,
      controlId: block.controlId,
      drawingNodeId: drawings[0]!.id,
      before: payload,
      resource: atom.resource,
    });
    // The original source coordinates still address this retained control in the recursive planner.
    replacements.set(old.path.join('/'), { ...next.ast, position: old.ast.position });
  }
  const visit = (node: TargetAstNode, path: readonly number[]): TargetAstNode =>
    replacements.get(path.join('/')) ??
    (node.children
      ? {
          ...node,
          children: (node.children as TargetAstNode[]).map((child, index) =>
            visit(child, [...path, index])
          ),
        }
      : node);
  const checked = createEditPlan(context.source.index.revision, edits);
  return checked.success
    ? { ast: visit(before.ast, []), edits: checked.plan.operations, alignment }
    : checked;
}
