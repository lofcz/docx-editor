import type { TreeDocxSessionView } from '@docx-editor.dev/core/binding';
import type { SemanticSelection } from '@docx-editor.dev/core/layout';
import type { StoryScope } from '@docx-editor.dev/core/store';
import { createBlockReservationState } from '../layout/block-reservations.ts';
import { createContentControlCollapseState } from '../layout/content-control-collapse-state.ts';
import type { SurfaceBlockReservations } from './surface-block-reservations.ts';
import type { SurfaceContentControlCollapse } from './surface-content-control-collapse.ts';

interface SurfaceBlockViewDeps {
  scope(): StoryScope;
  destroyed(): boolean;
  resolveNodeId(id: string): string | undefined;
  invalidate(reason: string): void;
  flush(): void;
}

/** Bind local block views to one canonical session and its existing layout scheduler. */
export function createSurfaceBlockView(
  session: Pick<TreeDocxSessionView, 'currentPackage' | 'paragraphIdsIn'>,
  deps: SurfaceBlockViewDeps
) {
  const state = createContentControlCollapseState();
  const reservations = createBlockReservationState();
  const part = () => {
    const pkg = session.currentPackage();
    return pkg.parts.get(pkg.mainDocumentPart)!;
  };
  const reveal = (paragraphIds: readonly string[]) => {
    if (!state.reveal(part(), paragraphIds)) return;
    deps.invalidate('content-control-collapse');
    deps.flush();
  };
  const commands: SurfaceContentControlCollapse & SurfaceBlockReservations = {
    collapsedContentControls: state.current,
    setCollapsedContentControls(next) {
      if (deps.destroyed()) return { ok: false, reason: 'document-not-ready' };
      const result = state.set(part(), next);
      if (!result.ok) return result;
      if (result.changed) deps.invalidate('content-control-collapse');
      return { ok: true };
    },
    setBlockReservations(next) {
      if (!deps.destroyed() && reservations.set(next)) deps.invalidate('block-reservations');
    },
    setBlockInsets(next) {
      if (!deps.destroyed() && reservations.setInsets(next)) deps.invalidate('block-reservations');
    },
  };
  return {
    commands,
    collapseOption: state.option,
    reservations: reservations.option,
    reconcile: () => state.reconcile(part(), deps.resolveNodeId),
    reveal,
    // Use before edits and after their selection moves: Enter can create a hidden body paragraph.
    revealSelection(next: SemanticSelection) {
      if (!state.current().length) return;
      const { anchor, head } = next;
      if (anchor.paragraphId === head.paragraphId) return reveal([head.paragraphId]);
      const order = session.paragraphIdsIn(deps.scope());
      const a = order.indexOf(anchor.paragraphId),
        b = order.indexOf(head.paragraphId);
      if (a >= 0 && b >= 0) reveal(order.slice(Math.min(a, b), Math.max(a, b) + 1));
    },
  };
}
