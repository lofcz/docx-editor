import type { PaginatedSurface } from './paginated-surface-contract.ts';
import type { TreeApplyResult } from '@docx-editor.dev/core/binding';
import type { SelectionMark } from '@docx-editor.dev/core/store';
import type { StoryScope } from '../store/store/tree-package-store.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';

/** The surface with the internal members that other editor modules reach by cast. */
export type ScaleMutableSurface = PaginatedSurface & {
  setScale(nextScale: number): boolean;
  /**
   * Why the mode rules would refuse these automation ops now, or null: the check
   * `applyAutomationOps` makes, without writing. A preview fork asks it so it refuses what a
   * commit would.
   */
  automationWriteRefusal(ops: readonly TreeDocOp[], packageEdits: boolean): string | null;
  /**
   * The gated, attributed session write for command lanes that live OUTSIDE this file —
   * today the content-control commands. Same collaboration gate and actor attribution as
   * the typing lane (`applyJournaledOps`); the command's own mode/lock gate has already
   * run. Internal, reached by cast like `setScale` — see `applyGatedSurfaceTreeOps` in
   * `content-controls.ts` for the accessor.
   */
  applyGatedTreeOps(
    ops: readonly TreeDocOp[],
    selectionBefore?: SelectionMark | null,
    selectionAfter?: SelectionMark | null,
    scope?: StoryScope
  ): TreeApplyResult;
};
