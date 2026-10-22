import type { EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { restoreDecisionEmptiedControls } from './content-control-prompt-restore.ts';
import { normalizeResolvedRunOrigins } from './run-origin-resolution.ts';

/** Restore native control structure before reconciling the fork's retained run origins. */
export function finishRevisionDecision(
  before: OoxmlPart,
  decided: ReadonlySet<string>,
  settled: OoxmlPart,
  origins: Parameters<typeof normalizeResolvedRunOrigins>[1],
  options?: EditOptions
) {
  return normalizeResolvedRunOrigins(
    restoreDecisionEmptiedControls(before, decided, settled, options),
    origins,
    options
  );
}
