import { trackHyperlinkStructurePlan } from './track-hyperlink-structure';
import { withReviewGroupRevisionReservations } from './group-allocation';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  collectRevisionSites,
  ensureHyperlinkRelationship,
  findNode,
  hyperlinkTargetOf,
  parentNodeOf,
  relationshipTargetIn,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
  type RevisionAddress,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { CanonicalEditPlan } from '../diff/plan';
import {
  nativePropertyAddress,
  propertyAddressKey,
  propertyControlRefusal,
} from './property-fallback';
import type { TrackTextContext } from './tracked-inline';
import { fail, readTrackedStructureInput } from './tracked-structure';
import { persistReviewGroups } from './groups';

/** The struck link and its retargeted copy; Word addresses survive save/reopen. */
export interface NativeHyperlinkProposal {
  readonly identity: string;
  readonly partName: string;
  readonly kind: 'hyperlinkTarget';
  /** The original link, which keeps its old target until the decision. */
  readonly hyperlinkId: string;
  readonly addresses: readonly RevisionAddress[];
}
/** One compound link intention, addressed by its actual persisted review group. */
export interface NativeHyperlinkStructureProposal {
  readonly identity: string;
  readonly partName: string;
  readonly kind: 'hyperlinkStructure';
  readonly addresses: readonly RevisionAddress[];
}
export type TrackHyperlinksResult =
  | {
      readonly success: true;
      readonly revision: RevisionToken;
      readonly proposals: readonly (NativeHyperlinkProposal | NativeHyperlinkStructureProposal)[];
    }
  | NativeEditFailure;

interface Retarget {
  readonly linkId: string;
  readonly url?: string;
  readonly relationshipId?: string;
  readonly anchor?: string;
  readonly tooltip?: string;
}

/**
 * Execute hyperlink target, tooltip, and D27 structure plans in one admitted native transaction.
 * Structure plans share the later-wave coordinator. Old values and the revision are rechecked
 * first; text and marks stay as they are in both views. Relationships are minted inside the
 * transaction and never rewritten, so a link sharing the old target keeps it.
 */
export function trackHyperlinkPlan(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext
): TrackHyperlinksResult {
  return withReviewGroupRevisionReservations(editor.surface?.session.currentPackage() ?? null, () =>
    trackHyperlinkPlanReserved(editor, input, context)
  );
}

function trackHyperlinkPlanReserved(
  editor: DocxEditorInstance,
  input: CanonicalEditPlan,
  context: TrackTextContext
): TrackHyperlinksResult {
  const read = readTrackedStructureInput(editor, input, context, 'hyperlink');
  if (!read.success) return read;
  const { plan, part, revision, surface, actual } = read.input;
  if (plan.operations.some((edit) => edit.kind === 'link' || edit.kind === 'unlink'))
    return trackHyperlinkStructurePlan(editor, plan, context);
  const pkg = surface.session.currentPackage();
  const retargets: Retarget[] = [];
  for (const edit of plan.operations) {
    if (edit.partName !== part.name)
      return fail('Tracked hyperlinks must belong to the body part.', 'unsupported-operation');
    if (edit.kind !== 'format' || edit.target.kind !== 'hyperlink')
      return fail('This executor accepts hyperlink target plans only.', 'unsupported-operation');
    const linkId = edit.target.hyperlinkId;
    const link = findNode(part, linkId);
    if (link?.kind !== 'hyperlink')
      return fail('The planned hyperlink no longer exists.', 'stale-revision');
    if (retargets.some((entry) => entry.linkId === linkId))
      return fail('Plan one tracked change per hyperlink.', 'unsupported-operation');
    const protection = propertyControlRefusal(part, linkId);
    if (protection) return fail(protection, 'unsupported-operation');
    const current = hyperlinkTargetOf(link as Extract<OoxmlNode, { kind: 'hyperlink' }>, (id) =>
      relationshipTargetIn(pkg, part.name, id)
    );
    let next: Retarget = {
      linkId,
      ...(current.relationshipId !== undefined ? { relationshipId: current.relationshipId } : {}),
      ...(current.anchor !== undefined ? { anchor: current.anchor } : {}),
    };
    for (const change of edit.changes) {
      if (change.name === 'target') {
        const before = change.before.present
          ? (change.before.value as Record<string, unknown>)
          : null;
        if (!before || before.kind !== current.kind || before.authored !== current.authored)
          return fail('The planned hyperlink target no longer matches.', 'stale-revision');
        const after = change.after.present ? (change.after.value as Record<string, unknown>) : null;
        if (!after || typeof after.authored !== 'string')
          return fail(
            'A hyperlink target cannot be removed by a format plan.',
            'unsupported-operation'
          );
        next =
          after.kind === 'external'
            ? { linkId, url: after.authored, ...(next.tooltip ? { tooltip: next.tooltip } : {}) }
            : {
                linkId,
                anchor: after.authored,
                ...(next.tooltip ? { tooltip: next.tooltip } : {}),
              };
      } else if (change.name === 'tooltip') {
        if (
          change.before.present
            ? change.before.value !== current.tooltip
            : current.tooltip !== undefined
        )
          return fail('The planned hyperlink tooltip no longer matches.', 'stale-revision');
        if (!change.after.present)
          return fail(
            'Removing a hyperlink tooltip needs native link structure planning (D27).',
            'unsupported-operation'
          );
        next = { ...next, tooltip: String(change.after.value) };
      } else return fail('Unknown native hyperlink change.', 'unsupported-operation');
    }
    retargets.push(next);
  }
  if (!retargets.length) return { success: true, revision: Object.freeze(actual), proposals: [] };
  const stage = (relate: (url: string) => string | null): TreeDocOp[] | null => {
    const ops: TreeDocOp[] = [];
    for (const entry of retargets) {
      const relationshipId =
        entry.url !== undefined ? (relate(entry.url) ?? undefined) : entry.relationshipId;
      if (entry.url !== undefined && !relationshipId) return null;
      ops.push({
        op: 'setHyperlinkTarget',
        linkId: entry.linkId,
        ...(relationshipId !== undefined ? { relationshipId } : {}),
        ...(entry.anchor !== undefined ? { anchor: entry.anchor } : {}),
        ...(entry.tooltip !== undefined ? { tooltip: entry.tooltip } : {}),
        revision,
      });
    }
    return ops;
  };
  // Immutable preview with scratch relationship mints: every retarget and both decisions.
  const originalSites = new Set(collectRevisionSites(part).map((site) => site.node.id));
  let scratch: OoxmlPackage = pkg;
  const previewOps = stage((url) => {
    const made = ensureHyperlinkRelationship(scratch, url, part.name);
    if (!made) return null;
    scratch = made.pkg;
    return made.relationshipId;
  });
  if (!previewOps) return fail('The native engine refuses this URL.', 'unsupported-operation');
  let preview = part;
  for (const op of previewOps) {
    const result = applyTreeOp(preview, op);
    if (!result.ok)
      return fail(`Native hyperlink proposal refused: ${result.reason}.`, 'unsupported-operation');
    preview = result.part;
  }
  const fresh = collectRevisionSites(preview).filter((site) => !originalSites.has(site.node.id));
  for (const action of ['acceptAllRevisions', 'rejectAllRevisions'] as const) {
    const decided = applyTreeOp(preview, {
      op: action,
      siteNodeIds: fresh.map((site) => site.node.id),
    });
    if (!decided.ok)
      return fail(
        `Native hyperlink ${action} preflight refused: ${decided.reason}.`,
        'unsupported-operation'
      );
  }
  let proposals: readonly NativeHyperlinkProposal[] = [];
  const complete = (working: OoxmlPackage): OoxmlPackage => {
    const current = working.parts.get(part.name)!;
    const byLink = retargets.map(() => new Map<string, RevisionAddress>());
    // The struck half sits inside the original link; the inserted half inside its copy, the
    // hyperlink that directly follows it.
    const owners = retargets.map((entry) => {
      const original = findNode(current, entry.linkId);
      const paragraph = original && parentNodeOf(current, entry.linkId);
      if (!original || !paragraph) return [];
      const index = paragraph.children.findIndex((child) => child.id === entry.linkId);
      const copy = paragraph.children[index + 1];
      return copy?.kind === 'hyperlink' ? [original, copy] : [original];
    });
    const holds = (node: OoxmlNode, id: string): boolean =>
      node.id === id ||
      (node.kind !== 'textValue' && node.children.some((child) => holds(child, id)));
    for (const site of collectRevisionSites(current)) {
      if (originalSites.has(site.node.id)) continue;
      const index = owners.findIndex((nodes) => nodes.some((node) => holds(node, site.node.id)));
      const address = nativePropertyAddress(site.node);
      if (index < 0 || !address) throw new Error('A new native revision escaped its planned link.');
      byLink[index]!.set(propertyAddressKey(part.name, address), Object.freeze({ ...address }));
    }
    proposals = Object.freeze(
      retargets.flatMap((entry, index) => {
        const group = [...byLink[index]!.values()];
        // A retarget naming the current target and tooltip proposes nothing.
        if (!group.length) return [];
        return [
          Object.freeze({
            identity: JSON.stringify([
              part.name,
              group.map((address) => propertyAddressKey(part.name, address)).sort(),
            ]),
            partName: part.name,
            kind: 'hyperlinkTarget' as const,
            hyperlinkId: entry.linkId,
            addresses: Object.freeze(group),
          }),
        ];
      })
    );
    return persistReviewGroups(
      pkg,
      working,
      proposals,
      editor.snapshot().collaborationStatus === 'inactive'
    );
  };
  try {
    complete(withPart(scratch, preview));
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error), 'unsupported-operation');
  }
  proposals = [];
  const session = surface.session;
  try {
    const result = surface.applyAutomationOps(
      stage,
      { kind: 'body' },
      [complete],
      true,
      true,
      'after'
    );
    if (result.rejected || !result.committed)
      return fail(
        String(result.reason ?? 'Native hyperlink changes did not commit.'),
        result.reason === 'review-module-required' ? 'unsupported-operation' : 'apply-failed'
      );
  } catch (error) {
    const changed = session.packageRevision() !== actual.revision;
    return {
      success: false,
      code: 'apply-failed',
      error: `${changed ? 'The document changed during a failed native call. Re-read before retrying. ' : ''}${error instanceof Error ? error.message : String(error)}`,
      expectedRevision: actual,
      actualRevision: { ...actual, revision: session.packageRevision() },
    };
  }
  return {
    success: true,
    revision: Object.freeze({ ...actual, revision: session.packageRevision() }),
    proposals,
  };
}
