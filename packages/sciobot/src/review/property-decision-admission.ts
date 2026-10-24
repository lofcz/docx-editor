import { TreePackageStore, type OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeReviewGroup } from './groups';

/** Withdrawal replaces a native reject, so it must first pass the same native gates. */
export function propertyDecisionRefusal(
  pkg: OoxmlPackage,
  groups: readonly NativeReviewGroup[]
): string | null {
  if (!groups.length) return null;
  const main = pkg.parts.get(pkg.mainDocumentPart);
  if (!main) return 'A review document body is missing.';
  const sites = new Map<string, Set<string>>();
  for (const group of groups)
    for (const member of group.native) {
      const ids = sites.get(member.partName) ?? new Set<string>();
      for (const id of member.siteNodeIds) ids.add(id);
      sites.set(member.partName, ids);
    }
  // A detached package transaction also checks settings.xml protection and cross-part admission.
  // Its restored snapshots never become live document content or ownership history.
  const preview = new TreePackageStore(pkg, main);
  const result = preview.transact(
    { kind: 'body' },
    (ctx) => {
      for (const [partName, ids] of sites)
        if (!ctx.applyTo(partName, { op: 'rejectAllRevisions', siteNodeIds: [...ids] })) return;
    },
    { recordsHistory: false }
  );
  return result.ok ? null : String(result.reason);
}
