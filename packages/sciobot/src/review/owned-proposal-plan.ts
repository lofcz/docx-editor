import {
  applyTreeOp,
  collectRevisionSites,
  canonicalOoxmlFingerprint,
  type OoxmlPackage,
  type RevisionAddress,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { EDIT_PLAN_LIMITS } from '../diff/plan';
import { GROUP_ID_LIMIT } from './group-markers';
import {
  readReviewGroups,
  type NativeReviewGroup,
  type ReviewGroupPropertyMember,
  type ReviewGroupReceipt,
} from './groups';
import { ATOMIC_REVIEW_EXECUTOR } from './property-record-schema';
import { readPropertyReviewStore } from './property-record-store';
import { decodeBlockPresenceRecord } from './block-presence-schema';
import { preparePresenceDecision } from './prepare-presence-decision';
import { preparePayloadProposalDecision } from './prepare-payload-decision';
import { propertyAddressKey } from './property-fallback';
import { ownedProposalSelectionProblem } from './owned-proposal-input';

/**
 * Host-retained ownership, not ownership inferred from an author name or document proximity.
 * Obtain native addresses from the host's proposal receipts and retain the resulting group ID.
 * Property ownership names the exact proposal/block pair. No canonical node IDs are trusted input.
 */
export interface OwnedProposalGroup {
  readonly groupId: string;
  readonly native: readonly ReviewGroupReceipt[];
  readonly properties: readonly ReviewGroupPropertyMember[];
}

export type OwnedProposalInverse = Extract<
  TreeDocOp,
  { op: 'rejectAllRevisions' | 'removeContentControl' }
>;

/** A revision-bound inverse plan. It contains addresses and intents, never canonical snapshots. */
export interface OwnedProposalRemovalPlan {
  readonly schemaVersion: 1;
  readonly revision: RevisionToken;
  readonly selection: readonly OwnedProposalGroup[];
  readonly partOps: readonly {
    readonly partName: string;
    readonly ops: readonly OwnedProposalInverse[];
  }[];
  /** R25 must append ordinary immutable rejection receipts in the same transaction. */
  readonly propertyDecisions: readonly (ReviewGroupPropertyMember & {
    readonly decision: 'reject';
  })[];
  /** Imported owned groups need their existing addresses retained before later ID allocation. */
  readonly retainNativeGroups: readonly {
    readonly groupId: string;
    readonly members: readonly { readonly partName: string; readonly revision: RevisionAddress }[];
  }[];
}

export type OwnedProposalRemovalResult =
  | { readonly success: true; readonly plan: OwnedProposalRemovalPlan }
  | NativeEditFailure;

const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const object = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const keys = (value: object, allowed: readonly string[]) =>
  Object.keys(value).every((key) => allowed.includes(key));
const id = (value: unknown): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= GROUP_ID_LIMIT;
const propertyKey = (member: ReviewGroupPropertyMember) =>
  JSON.stringify([member.proposalId, member.blockId]);
const siteKey = (part: string, nodeId: string) => JSON.stringify([part, nodeId]);

function matchesOwnership(group: NativeReviewGroup, owned: OwnedProposalGroup): boolean {
  const native = new Set(
      owned.native.flatMap((member) =>
        member.addresses.map((revision) => propertyAddressKey(member.partName, revision))
      )
    ),
    properties = new Set(owned.properties.map(propertyKey));
  return (
    native.size === group.native.length &&
    group.native.every((member) =>
      native.has(propertyAddressKey(member.partName, member.address))
    ) &&
    properties.size === group.properties.length &&
    group.properties.every((member) => properties.has(propertyKey(member)))
  );
}

/**
 * Plan rejection of exactly the selected owned pending groups. This does not mutate a package,
 * editor, display mode, or history. The caller supplies the package's current revision token.
 * R25 must re-read/re-plan inside the native gate and apply all operations and metadata together.
 */
export function planOwnedProposalRemoval(
  source: { readonly pkg: OoxmlPackage; readonly revision: RevisionToken },
  selection: readonly OwnedProposalGroup[]
): OwnedProposalRemovalResult {
  try {
    const copied = copyBoundedJson({ revision: source.revision, selection }, EDIT_PLAN_LIMITS);
    if (!copied.ok) return fail(copied.reason);
    const input = copied.value as unknown as {
      revision: RevisionToken;
      selection: readonly OwnedProposalGroup[];
    };
    const token = input.revision;
    if (
      !object(token) ||
      !id(token.documentId) ||
      !id(token.epoch) ||
      !Number.isSafeInteger(token.revision) ||
      token.revision < 0 ||
      !keys(token, ['documentId', 'epoch', 'revision'])
    )
      return fail('Invalid owned proposal revision.');
    const problem = ownedProposalSelectionProblem(input.selection);
    if (problem) return fail(problem);
    const read = readReviewGroups(source.pkg);
    if (!read.success) return read;
    const owned = new Map(input.selection.map((group) => [group.groupId, group]));
    const selected = read.groups.filter((group) => owned.has(group.id));
    if (selected.length !== owned.size)
      return fail('An owned pending group is missing or already resolved.', 'stale-revision');
    if (selected.some((group) => group.readOnly))
      return fail(
        'An owned group needs reconciliation before replacement.',
        'unsupported-operation'
      );
    if (selected.some((group) => !matchesOwnership(group, owned.get(group.id)!)))
      return fail('Owned receipts do not match the complete current group.', 'ambiguous-target');

    const presenceByPart = new Map<string, OwnedProposalInverse[]>();
    const sitesByPart = new Map<string, Set<string>>(),
      propertyDecisions: OwnedProposalRemovalPlan['propertyDecisions'][number][] = [],
      retainNativeGroups: OwnedProposalRemovalPlan['retainNativeGroups'][number][] = [];
    const proposals = readPropertyReviewStore(source.pkg);
    for (const group of selected) {
      const expectedPropertySites = new Map<string, Set<string>>();
      for (const member of group.properties) {
        const entry =
          proposals.status === 'ready' ? proposals.records.get(member.proposalId) : undefined;
        if (!entry) return fail('An owned application proposal is missing.', 'invalid-document');
        if (entry.decoded.status !== 'supported') {
          if (decodeBlockPresenceRecord(entry.raw).status !== 'supported')
            return fail('An owned application proposal is unsupported.', 'unsupported-operation');
          const prepared = preparePresenceDecision(
            source.pkg,
            member.proposalId,
            member.blockId,
            'reject'
          );
          if (!prepared.success) return prepared;
          if (prepared.nativeSiteNodeIds) {
            const expected = expectedPropertySites.get(prepared.partName) ?? new Set<string>();
            for (const id of prepared.nativeSiteNodeIds) expected.add(id);
            expectedPropertySites.set(prepared.partName, expected);
          }
          const ops = presenceByPart.get(prepared.partName) ?? [];
          for (const op of prepared.ops) {
            if (op.op !== 'removeContentControl')
              return fail('An owned presence inverse has an unsupported native operation.');
            ops.push(op);
          }
          if (ops.length) presenceByPart.set(prepared.partName, ops);
          propertyDecisions.push({ ...member, decision: 'reject' });
          continue;
        }
        const prepared = preparePayloadProposalDecision(
          source.pkg,
          token,
          member.proposalId,
          'reject',
          entry.decoded.value.extension?.executor === ATOMIC_REVIEW_EXECUTOR
            ? 'atomic'
            : 'properties',
          new Set([member.blockId])
        );
        if (!prepared.success) return prepared;
        const ids = expectedPropertySites.get(prepared.partName) ?? new Set<string>();
        for (const op of prepared.ops)
          if (op.op === 'rejectRevision')
            for (const nodeId of op.siteNodeIds ?? []) ids.add(nodeId);
        expectedPropertySites.set(prepared.partName, ids);
        propertyDecisions.push({ ...member, decision: 'reject' });
      }
      if (group.properties.length) {
        const actual = new Set(
          group.native.flatMap((member) =>
            member.siteNodeIds.map((nodeId) => siteKey(member.partName, nodeId))
          )
        );
        const expected = new Set(
          [...expectedPropertySites].flatMap(([part, ids]) =>
            [...ids].map((nodeId) => siteKey(part, nodeId))
          )
        );
        if (actual.size !== expected.size || [...actual].some((key) => !expected.has(key)))
          return fail(
            'An owned property group has incomplete native membership.',
            'invalid-document'
          );
      } else if (!group.id.startsWith('g1:stored:')) {
        retainNativeGroups.push({
          groupId: group.id,
          members: group.native.map((member) => ({
            partName: member.partName,
            revision: member.address,
          })),
        });
      }
      for (const member of group.native) {
        const ids = sitesByPart.get(member.partName) ?? new Set<string>();
        for (const nodeId of member.siteNodeIds) ids.add(nodeId);
        sitesByPart.set(member.partName, ids);
      }
    }

    const partOps: OwnedProposalRemovalPlan['partOps'][number][] = [];
    for (const [name, ids] of sitesByPart) {
      const part = source.pkg.parts.get(name);
      if (!part) return fail('An owned revision part is missing.', 'invalid-document');
      const op: OwnedProposalInverse = { op: 'rejectAllRevisions', siteNodeIds: [...ids] };
      const preview = applyTreeOp(part, op);
      if (!preview.ok)
        return fail(`Owned revision inverse refused: ${preview.reason}.`, 'unsupported-operation');
      const before = collectRevisionSites(part),
        beforeIds = new Set(before.map((site) => site.node.id)),
        after = new Map(collectRevisionSites(preview.part).map((site) => [site.node.id, site]));
      if (before.some((site) => ids.has(site.node.id) && after.has(site.node.id)))
        return fail(
          'An owned revision inverse leaves selected sites pending.',
          'unsupported-operation'
        );
      for (const site of before) {
        if (ids.has(site.node.id)) continue;
        const retained = after.get(site.node.id);
        // An owned edit can sit inside a foreign wrapper. Its children may change through that
        // exact inverse, but the foreign revision's identity and authored attributes must survive.
        if (
          !retained ||
          retained.node.kind !== site.node.kind ||
          retained.node.namespaceUri !== site.node.namespaceUri ||
          retained.node.localName !== site.node.localName ||
          JSON.stringify(retained.node.attributes) !== JSON.stringify(site.node.attributes)
        )
          return fail('Owned removal would change an unselected revision.', 'ambiguous-target');
      }
      if ([...after.keys()].some((nodeId) => !beforeIds.has(nodeId)))
        return fail('Owned removal would introduce an unrelated revision.', 'ambiguous-target');
      partOps.push({ partName: name, ops: [op] });
    }
    for (const [partName, ops] of presenceByPart) {
      let part = source.pkg.parts.get(partName);
      if (!part) return fail('An owned presence part is missing.', 'invalid-document');
      const existing = partOps.find((row) => row.partName === partName);
      for (const op of existing?.ops ?? []) {
        const applied = applyTreeOp(part, op);
        if (!applied.ok) return fail('An owned native inverse refused.', 'unsupported-operation');
        part = applied.part;
      }
      for (const op of ops) {
        const before = collectRevisionSites(part);
        const applied = applyTreeOp(part, op);
        if (!applied.ok) return fail('An owned presence inverse refused.', 'unsupported-operation');
        const after = new Map(
          collectRevisionSites(applied.part).map((site) => [site.node.id, site.node])
        );
        if (
          before.length !== after.size ||
          before.some((site) => {
            const retained = after.get(site.node.id);
            return (
              !retained ||
              canonicalOoxmlFingerprint(retained) !== canonicalOoxmlFingerprint(site.node)
            );
          })
        )
          return fail(
            'Owned presence removal would change an unselected revision.',
            'ambiguous-target'
          );
        part = applied.part;
      }
      if (existing)
        partOps[partOps.indexOf(existing)] = { partName, ops: [...existing.ops, ...ops] };
      else partOps.push({ partName, ops });
    }
    const plan: OwnedProposalRemovalPlan = {
      schemaVersion: 1,
      revision: token,
      selection: selected.map((group) => owned.get(group.id)!),
      partOps,
      propertyDecisions,
      retainNativeGroups,
    };
    const bounded = copyBoundedJson(plan, EDIT_PLAN_LIMITS);
    if (!bounded.ok) return fail(bounded.reason);
    freeze(bounded.value);
    return { success: true, plan: bounded.value as unknown as OwnedProposalRemovalPlan };
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause), 'invalid-document');
  }
}

function freeze(value: unknown): void {
  if (value === null || typeof value !== 'object') return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}
