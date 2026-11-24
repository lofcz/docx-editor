import type { CollapsedContentControl } from '../contracts/content-control-collapse.ts';
import type { LayoutScope } from '../layout/layout-scheduler.ts';
import { snapshotLayoutSession } from '../layout/layout-session-snapshot.ts';
import type { LayoutSession } from '../layout/layout-session.ts';

type Folds = readonly CollapsedContentControl[];

interface FoldLayout {
  revision: number;
  folds: Folds;
  preceding?: { folds: ReadonlyMap<string, CollapsedContentControl>; restore(): void };
}

/** Remember the preceding local disclosure's flow checkpoints for its return trip. */
export function createDisclosureLayoutHistory(readFolds: () => Folds) {
  const histories = new WeakMap<LayoutSession, FoldLayout>();
  return (session: LayoutSession, scope: LayoutScope | undefined, settled: boolean) => {
    if (!scope) {
      const revision = session.previous?.revision;
      if (revision !== undefined && !histories.has(session))
        histories.set(session, { revision, folds: readFolds() });
      return;
    }
    // The first view of an inline widget can supply its size after disclosure.
    // The next fold can save this newly measured layout, but the preceding
    // checkpoint still belongs to the old measurements.
    const unchangedDocument =
      settled && !scope.paragraphIds.size && !scope.created.size && !scope.deleted.size;
    if (
      unchangedDocument &&
      scope.dependencyKeys.size === 1 &&
      scope.dependencyKeys.has('inline-widgets') &&
      histories.has(session)
    ) {
      histories.set(session, { revision: scope.revision, folds: readFolds() });
      return;
    }
    if (
      !unchangedDocument ||
      !scope.dependencyKeys.has('content-control-collapse') ||
      [...scope.dependencyKeys].some(
        (key) => key !== 'content-control-collapse' && key !== 'block-reservations'
      )
    ) {
      histories.delete(session);
      return;
    }
    const folds = readFolds();
    const previous = histories.get(session);
    const current: FoldLayout = { revision: scope.revision, folds };
    histories.set(session, current);
    if (previous?.revision !== scope.revision || session.previous?.revision !== scope.revision)
      return;
    current.preceding = {
      folds: new Map(previous.folds.map((fold) => [fold.controlId, fold])),
      restore: snapshotLayoutSession(session),
    };
    const saved = previous.preceding;
    if (saved?.folds.size !== folds.length) return;
    for (const fold of folds) {
      const other = saved.folds.get(fold.controlId);
      if (
        !other ||
        other.summaryParagraphId !== fold.summaryParagraphId ||
        other.placeholderHeightPt !== fold.placeholderHeightPt
      )
        return;
    }
    // The following normal pass still checks current fonts, widget sizes, section geometry,
    // fields and flow keys. Reuse selects a starting checkpoint, not a publishable result.
    const fullPasses = session.stats.fullPasses;
    saved.restore();
    session.stats = { ...session.stats, fullPasses };
  };
}
