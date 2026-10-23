import {
  collectRevisionSites,
  paragraphTextOf,
  prepareTrackedRunOrigins,
  withPart,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit, CanonicalEditPlan } from '../diff/plan';
import { leadingReplacementEdit } from './leading-replacements';
import { prepareTrackedInlinePlan } from './tracked-inline';
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
  revision: TrackedStageAttribution
): StagedTextCoordinates {
  const first = edits[0];
  if (!first) refuseWrite('A text stage needs one paragraph.');
  const before = writer.currentPackage().parts.get(first.partName);
  if (!before) refuseWrite('The native text story disappeared.', 'invalid-document');
  const original = paragraphTextOf(before, first.paragraphId);
  if (original === null || edits.some((edit) => edit.paragraphId !== first.paragraphId))
    refuseWrite('A text stage needs one native paragraph.', 'invalid-input');
  const prepared = prepareTrackedInlinePlan(
    before,
    { ...plan, operations: edits },
    revision,
    revision.date
  );
  if (!prepared.success) refuseWrite(prepared.error, prepared.code);
  let runOriginScope: string | undefined;
  writer.edit((pkg) => {
    const marked = prepareTrackedRunOrigins(pkg.parts.get(before.name)!, first.paragraphId, edits);
    runOriginScope = marked.scope ?? undefined;
    return withPart(pkg, marked.part);
  });
  const oldSites = new Set(collectRevisionSites(before).map((site) => site.node.id));
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
