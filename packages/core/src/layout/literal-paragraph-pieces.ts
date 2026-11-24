import type { DisplayWalkInput } from './field-display-input.ts';
import { withRunFormatSite, type FieldAwarePiece } from './field-pieces.ts';
import { applyParagraphFontSlots } from './field-projection-finish.ts';
import { modelTextOfRunChild, runPropertiesOf } from './field-run-text.ts';
import { resolvedFormatChangeOf } from './revision-formatting-projection.ts';
import {
  DEFAULT_REVISION_DISPLAY_MODE,
  projectedRevisionProperties,
} from './revision-projection.ts';
import { resolveRunStyle } from './run-style.ts';

/** Direct text runs have no field state, wrappers, or projected atoms. */
export function literalParagraphPieces(input: DisplayWalkInput): FieldAwarePiece[] | undefined {
  const paragraph = input.paragraph;
  if (
    paragraph.kind !== 'paragraph' ||
    (input.authorFilter?.revisionMarkup &&
      (input.displayMode ?? DEFAULT_REVISION_DISPLAY_MODE) === 'all-markup') ||
    input.showFieldCodes ||
    input.tocLinkStyleRanges?.length ||
    !paragraph.children.every(
      (child) =>
        child.kind === 'paragraphProperties' ||
        (child.kind === 'run' &&
          child.children.every((item) => item.kind === 'text' || item.kind === 'runProperties'))
    )
  )
    return undefined;
  const pieces: FieldAwarePiece[] = [];
  let offset = 0;
  for (const run of paragraph.children) {
    if (run.kind !== 'run') continue;
    const props = runPropertiesOf(run, input.inheritedRunProperties ?? [], input.cascadeRuns);
    const style = resolveRunStyle(props, input.themeFonts);
    const site = resolvedFormatChangeOf(
      run.children.find((child) => child.kind === 'runProperties'),
      input.authorFilter
    );
    const attribution = withRunFormatSite(
      { props: projectedRevisionProperties(props, input.authorFilter) },
      site
    );
    for (const child of run.children) {
      if (child.kind !== 'text') continue;
      const text = modelTextOfRunChild(child);
      const start = offset;
      offset += text.length;
      if (text && !style.hidden) pieces.push({ text, ...attribution, style, start, end: offset });
    }
  }
  return applyParagraphFontSlots(pieces, input.themeFonts);
}
