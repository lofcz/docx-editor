import {
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  withPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

type Insertion = Omit<Extract<TreeDocOp, { op: 'insertInlineContentControl' }>, 'revision'>;

/** Retain original run boundaries before native insertion; a prior text stage can lend its scope. */
export function stageTrackedInlineControl(
  writer: TrackedStageWriter,
  partName: string,
  command: Insertion,
  attribution: TrackedStageAttribution,
  suppliedScope?: string
): void {
  const before = writer.currentPackage().parts.get(partName);
  if (!before) refuseWrite('The inline proposal story is missing.', 'invalid-document');
  let scope = suppliedScope ?? null;
  if (!scope)
    writer.edit((pkg) => {
      const prepared = prepareTrackedRunOrigins(pkg.parts.get(partName)!, command.paragraphId, [
        { start: command.offset, end: command.offset },
      ]);
      scope = prepared.scope;
      return withPart(pkg, prepared.part);
    });
  writer.apply(partName, { ...command, revision: attribution });
  if (scope)
    writer.edit((pkg) =>
      withPart(
        pkg,
        sealTrackedRunOrigins(pkg.parts.get(partName)!, command.paragraphId, scope!, before!)
      )
    );
}
