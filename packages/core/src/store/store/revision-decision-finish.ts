import type { EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlPart, OoxmlNode } from '../package/ooxml-tree.ts';
import { restoreDecisionEmptiedControls } from './content-control-prompt-restore.ts';
import { normalizeResolvedRunOrigins } from './run-origin-resolution.ts';
import { readRevisionOrigins } from './run-origin-metadata.ts';

/** Restore native control structure before reconciling the fork's retained run origins. */
export function finishRevisionDecision(
  before: OoxmlPart,
  decided: ReadonlySet<string>,
  settled: OoxmlPart,
  origins: Parameters<typeof normalizeResolvedRunOrigins>[1],
  revisionSites: (part: OoxmlPart) => readonly { readonly node: OoxmlNode }[],
  options?: EditOptions
) {
  const restored = restoreDecisionEmptiedControls(before, decided, settled, options);
  // The decision owns classification. Export readers must not import the decision executor.
  const pending = new Set(
    revisionSites(restored)
      .flatMap(({ node }) => readRevisionOrigins(node))
      .map(({ scope }) => scope)
  );
  return normalizeResolvedRunOrigins(restored, origins, pending, options);
}
