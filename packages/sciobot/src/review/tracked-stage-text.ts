import {
  collectRevisionSites,
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  prepareTrackedRunOrigins,
  withPart,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit, CanonicalEditPlan } from '../diff/plan';
import { leadingReplacementEdit } from './leading-replacements';
import { prepareTrackedInlinePlan } from './tracked-inline';
import { trackedTextCoordinateEdits } from './tracked-text-coordinate-edits';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';
import {
  stagedParagraphCoordinates,
  type StagedTextCoordinates,
} from './tracked-stage-coordinates';
export type { StagedTextCoordinates } from './tracked-stage-coordinates';

/** Apply one paragraph's original-coordinate text edits before its dependent format/link steps. */
export function stageTrackedText(
  writer: TrackedStageWriter,
  plan: CanonicalEditPlan,
  edits: readonly Extract<CanonicalEdit, { kind: 'text' }>[],
  revision: TrackedStageAttribution,
  dependent?: {
    readonly coordinates: StagedTextCoordinates;
    readonly initialSites: ReadonlySet<string>;
  }
): StagedTextCoordinates {
  const first = edits[0];
  if (!first) refuseWrite('A text stage needs one paragraph.');
  const before = writer.currentPackage().parts.get(first.partName);
  if (!before) refuseWrite('The native text story disappeared.', 'invalid-document');
  const original = dependent?.coordinates.text() ?? paragraphTextOf(before, first.paragraphId);
  if (original === null || edits.some((edit) => edit.paragraphId !== first.paragraphId))
    refuseWrite('A text stage needs one native paragraph.', 'invalid-input');
  const nativeEdits = dependent ? trackedTextCoordinateEdits(edits, dependent.coordinates) : edits;
  const prepared = prepareTrackedInlinePlan(
    before,
    { ...plan, operations: nativeEdits },
    revision,
    revision.date
  );
  if (!prepared.success) refuseWrite(prepared.error, prepared.code);
  const paragraph = findNode(before, first.paragraphId);
  if (paragraph?.kind !== 'paragraph')
    refuseWrite('The native text paragraph disappeared.', 'invalid-document');
  const offsets = paragraphOffsetIndex(paragraph);
  for (const site of collectRevisionSites(before)) {
    if (site.node.kind !== 'revisionInsert' && site.node.kind !== 'revisionDelete') continue;
    const span = offsets.spanOf(site.node);
    if (
      span &&
      nativeEdits.some((edit) =>
        edit.start === edit.end
          ? span.start < edit.start && edit.start < span.end
          : span.start < edit.end && edit.start < span.end
      )
    )
      refuseWrite('Replacement would change an unselected revision.', 'ambiguous-target');
  }
  let runOriginScope: string | undefined;
  writer.edit((pkg) => {
    const marked = prepareTrackedRunOrigins(
      pkg.parts.get(before.name)!,
      first.paragraphId,
      nativeEdits
    );
    runOriginScope = marked.scope ?? undefined;
    return withPart(pkg, marked.part);
  });
  const oldSites =
    dependent?.initialSites ?? new Set(collectRevisionSites(before).map((site) => site.node.id));
  for (const op of prepared.ops) writer.apply(before.name, op);
  writer.edit(
    leadingReplacementEdit(before, prepared.paragraphIds, revision.author, revision.date)
  );
  let expected = original;
  for (const edit of [...edits].sort((a, b) => b.start - a.start))
    expected = expected.slice(0, edit.start) + edit.after + expected.slice(edit.end);
  return {
    ...stagedParagraphCoordinates(writer, before.name, first.paragraphId, oldSites, expected),
    ...(runOriginScope ? { runOriginScope } : {}),
  };
}

export function currentStagePart(writer: TrackedStageWriter, partName: string): OoxmlPart {
  const part = writer.currentPackage().parts.get(partName);
  if (!part) refuseWrite('The native proposal story disappeared.', 'invalid-document');
  return part;
}
