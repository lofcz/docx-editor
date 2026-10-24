import type { TreeApplyOptions, TreeApplyResult } from '@docx-editor.dev/core/binding';
import type { StoryScope, TreeDocOp, SelectionMark } from '@docx-editor.dev/core/store';
import type { SemanticSelection } from '@docx-editor.dev/core/layout';
import { runWithTransactionActor } from '../store/package/actor-scoped-ids.ts';
import { isTrackedEdit } from './surface-tracked-ops.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';

interface SurfaceAutomationDeps {
  commit(
    run: () => TreeApplyResult,
    selection: () => SemanticSelection | null,
    options: { automation: boolean }
  ): string | null;
  writeRefusal(edits: boolean, ops?: readonly TreeDocOp[], checkSelection?: boolean): string | null;
  reviewAvailable(): boolean;
  actor(): string | undefined;
  mintHyperlink(url: string, story: StoryScope): string | null;
  applyOps(
    ops: readonly TreeDocOp[],
    before?: SelectionMark | null,
    after?: SelectionMark | null,
    scope?: StoryScope,
    checkSelection?: boolean,
    packageEdits?: TreeApplyOptions['packageEdits'],
    packageEditsAfterOps?: TreeApplyOptions['packageEditsAfterOps'],
    stage?: TreeApplyOptions['stage']
  ): TreeApplyResult;
  selectionAfter(): SemanticSelection | null;
  unexplainedReason: string;
}

/** Automation and composite native gestures share the ordinary write gates and journal. */
export function createSurfaceAutomation(
  deps: SurfaceAutomationDeps
): PaginatedSurface['applyAutomationOps'] {
  const {
    commit,
    writeRefusal,
    reviewAvailable,
    actor,
    mintHyperlink,
    applyOps,
    selectionAfter,
    unexplainedReason,
  } = deps;
  return (
    staged,
    scope,
    packageEdits,
    requiresReview = false,
    separateHistory = false,
    packageEditPhase = 'before',
    stage
  ) => {
    // Automation shares edit gates and repainting. Its story and history boundaries remain explicit.
    let result: TreeApplyResult | null = null;
    const story = scope ?? { kind: 'body' };
    const refusal = commit(
      () => {
        // Check viewing before building ops can mint relationships outside the transaction and undo stack.
        // `edits: false` also permits revision decisions. `applyOps` checks each resulting operation.
        const viewing = writeRefusal(false);
        if (viewing !== null) {
          return (result = { committed: false, rejected: true, opCount: 0, reason: viewing });
        }
        // The mint carries the edit rule with it: it IS an edit, so a mode that would refuse one
        // must refuse it, and refusing here means the relationship is never written.
        let refused: string | null = null;
        const ops = staged((url) => {
          const refusal = writeRefusal(true, [], false);
          // Bound to the collaboration actor for the same reason the mint is called out above:
          // it lands on the package outside the transaction, so nothing else would bind one, and
          // two peers linking at once would otherwise agree on `rId${max + 1}`.
          if (refusal === null) {
            return runWithTransactionActor(actor(), () => mintHyperlink(url, story));
          }
          refused = refusal;
          return null;
        });
        if (ops === null) {
          return (result = {
            committed: false,
            rejected: true,
            opCount: 0,
            reason: refused ?? 'this engine will not author that hyperlink target',
          });
        }
        if (!reviewAvailable() && (requiresReview || ops.some(isTrackedEdit))) {
          return (result = {
            committed: false,
            rejected: true,
            opCount: 0,
            reason: 'review-module-required',
          });
        }
        return (result = applyOps(
          ops,
          undefined,
          undefined,
          story,
          false,
          packageEditPhase === 'before' ? packageEdits : undefined,
          packageEditPhase === 'after' ? packageEdits : undefined,
          stage
        ));
      },
      selectionAfter,
      { automation: !separateHistory }
    );
    const reason = refusal ?? unexplainedReason; // Null means the refresh guard skipped the thunk.
    return result ?? { committed: false, rejected: true, opCount: 0, reason };
  };
}
