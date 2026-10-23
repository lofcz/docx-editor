import type { OoxmlPart } from '@docx-editor.dev/core/store';
import { hyperlinkCharacterStyle, inlineCodeCharacterStyles } from '../native/inline-code';
import type { TargetParagraphSource } from './target-inline';
import type { CompiledInlineTarget } from './inline-target';
import { planMarkDeltas } from './insertion-marks';
import { createEditPlan, type EditPlanResult } from './plan';
import { retainedAtomStreams } from './retained-atom-stream';
import { planLinkStructure } from './target-links';

/** Format unchanged atoms and reconcile their link envelopes without replacing native controls. */
export function planRetainedAtomMarks(
  context: TargetParagraphSource,
  part: OoxmlPart,
  paragraphId: string,
  before: CompiledInlineTarget,
  after: CompiledInlineTarget,
  base: number
): EditPlanResult | null {
  const streams = retainedAtomStreams(context, part, paragraphId, before, after, base);
  if (!streams) return null;
  const runs = (target: CompiledInlineTarget) =>
    target.tokens.flatMap((token) =>
      token.kind === 'text' ? [{ value: token.value, marks: token.marks }] : []
    );
  const planned = planMarkDeltas(
    {
      revision: context.source.index.revision,
      part,
      paragraphId,
      base,
      length: streams.after.plainText!.length,
      displayMode: context.input.projection.displayMode!,
      hyperlinkStyle: hyperlinkCharacterStyle(context.input.pkg),
      codeStyles: inlineCodeCharacterStyles(context.input.pkg),
    },
    runs(streams.before),
    runs(streams.after)
  );
  if (!planned.success) return planned;
  const links = planLinkStructure(
    { revision: context.source.index.revision, pkg: context.input.pkg, part, paragraphId, base },
    streams.after
  );
  if (!links.success) return links;
  return createEditPlan(context.source.index.revision, [
    ...planned.operations,
    ...links.operations,
  ]);
}
