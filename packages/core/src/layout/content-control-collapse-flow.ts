import type { OoxmlElement, OoxmlPart } from '@docx-editor.dev/core/store';
import { visibleCollapsedControlFlow } from './content-control-collapse-state.ts';
import { noteRoundBlocks } from './note-reserve-search.ts';
import { enumerateDocumentSectionsFromBlocks, type DocumentSection } from './section-properties.ts';
import type { SemanticLayoutOptions } from './semantic-layout-options.ts';

/** Filter flow after canonical counters and references resolve, before note pagination rounds. */
export function visibleBodyRound(
  part: OoxmlPart,
  blocks: readonly OoxmlElement[],
  sections: readonly DocumentSection[],
  options: SemanticLayoutOptions
) {
  const visible = visibleCollapsedControlFlow(blocks, options.contentControlCollapse);
  // Validation excludes hidden section boundaries, so section properties retain their owners.
  const visibleSections =
    visible === blocks ? sections : enumerateDocumentSectionsFromBlocks(part, visible).sections;
  return noteRoundBlocks(part, visible, visibleSections, options);
}

/** Visibility invalidates placement checkpoints without invalidating paragraph measurements. */
export function bodyFlowProjectionEpoch(options: SemanticLayoutOptions): string | null {
  return options.projectionTokenForParagraph !== undefined && options.projectionEpoch === undefined
    ? null
    : (options.projectionEpoch ?? '') +
        (options.blockReservations?.epoch ?? '') +
        (options.contentControlCollapse?.token ?? '');
}
