import { decodeBlockPresenceRecord } from './block-presence-schema';
import { bindBlockPresenceRecord } from './block-presence-plan';
import {
  collectRevisionSites,
  isStoryPart,
  locateSites,
  revisionItemsOf,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type RevisionAddress,
  type ReviewRange,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import type { NativeEditFailure } from '../contracts/document';
import { readPropertyReviewStore } from './property-record-store';
import {
  atomicFallbackParagraph,
  nativePropertyAddress,
  propertyAddressKey,
  propertyFallbackParagraph,
} from './property-fallback';
import { reviewGroupMarker } from './group-markers';
import { readPayloadGroupDecisions, payloadGroupDecisionKey } from './payload-group-decisions';
import {
  readReviewGroupStore,
  withInitializedReviewGroupStore,
  withReviewGroupRecord,
} from './group-record-store';

export interface ReviewGroupNativeMember {
  readonly partName: string;
  readonly address: RevisionAddress;
  /** Exact sites in this revision only. Re-read before any decision. */
  readonly siteNodeIds: readonly string[];
}
export interface ReviewGroupPropertyMember {
  readonly proposalId: string;
  readonly blockId: string;
}
export interface ReviewGroupAnchor {
  readonly partName: string;
  readonly nodeId: string;
  readonly path: readonly number[];
  readonly paragraphId: string | null;
  readonly offset: number;
}
/** Stable identity; canonical site IDs and paths are a disposable read projection. */
export interface NativeReviewGroup {
  readonly id: string;
  readonly native: readonly ReviewGroupNativeMember[];
  readonly properties: readonly ReviewGroupPropertyMember[];
  readonly anchor: ReviewGroupAnchor | null;
  readonly readOnly: boolean;
}
export type ReviewGroupsResult =
  | { readonly success: true; readonly groups: readonly NativeReviewGroup[] }
  | NativeEditFailure;
export interface ReviewGroupReceipt {
  readonly partName: string;
  readonly addresses: readonly RevisionAddress[];
}
interface Site {
  readonly key: string;
  readonly partName: string;
  readonly node: OoxmlElement;
  readonly address: RevisionAddress | null;
  readonly path: readonly number[];
  readonly range: ReviewRange | null;
  readonly refused: boolean;
  readonly kind: string;
  readonly marker: string | undefined;
}
interface Bucket {
  id?: string;
  sites: Site[];
  properties: ReviewGroupPropertyMember[];
  anchor?: ReviewGroupAnchor;
  readOnly: boolean;
}
const keyOf = (part: string, nodeId: string) => JSON.stringify([part, nodeId]);
const pathCompare = (a: readonly number[], b: readonly number[]) => {
  for (let index = 0; index < Math.min(a.length, b.length); index++)
    if (a[index] !== b[index]) return a[index]! - b[index]!;
  return a.length - b.length;
};
const within = (path: readonly number[], parent: readonly number[]) =>
  parent.every((step, index) => path[index] === step);
const anchorOf = (site: Site): ReviewGroupAnchor => ({
  partName: site.partName,
  nodeId: site.node.id,
  path: site.path,
  paragraphId: site.range?.start.paragraphId ?? null,
  offset: site.range?.start.offset ?? 0,
});
const payloadGroupId = (proposalId: string, blockId: string) =>
  `g1:payload:${JSON.stringify([proposalId, blockId])}`;
function groupAddress(node: OoxmlElement): RevisionAddress | null {
  const address = nativePropertyAddress(node);
  if (address || node.localName !== 'tblGridChange') return address;
  // CT_TblGridChange has an ID and no author. Match the native table review address.
  const id = node.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'id'
  )?.value;
  return id === undefined ? null : { id, author: '' };
}

/** Native reading only. No display mode, document content, metadata, or history mutation. */
export function readReviewGroups(pkg: OoxmlPackage): ReviewGroupsResult {
  try {
    return { success: true, groups: derive(pkg).groups };
  } catch (cause) {
    return {
      success: false,
      code: 'invalid-document',
      error: cause instanceof Error ? cause.message : String(cause),
    };
  }
}

function derive(
  pkg: OoxmlPackage,
  receipts: readonly ReviewGroupReceipt[] = [],
  inherited: ReadonlyMap<string, { id: string; readOnly: boolean }> = new Map()
) {
  const parts = [...pkg.parts.values()]
    .filter((part) => isStoryPart(part) || collectRevisionSites(part).length > 0)
    .sort(
      (a, b) =>
        Number(b.name === pkg.mainDocumentPart) - Number(a.name === pkg.mainDocumentPart) ||
        a.name.localeCompare(b.name)
    );
  const partOrder = new Map(parts.map((part, index) => [part.name, index])),
    sites: Site[] = [],
    paths = new Map<string, readonly number[]>(),
    byKey = new Map<string, Site>(),
    owner = new Map<string, Bucket>(),
    buckets: Bucket[] = [],
    marked = new Map<string, Bucket>(),
    applicationScopes = new Map<
      Bucket,
      {
        partName: string;
        path: readonly number[];
        paragraphId: string | null;
        addresses: ReadonlySet<string>;
      }
    >(),
    byAddress = new Map<string, Site[]>();
  const itemsByPart = new Map(parts.map((part) => [part.name, revisionItemsOf(part)]));
  for (const part of parts) {
    const revisions = new Map(collectRevisionSites(part).map((site) => [site.node.id, site])),
      locations = locateSites(part),
      active = new Set(
        itemsByPart
          .get(part.name)!
          .flatMap((item) =>
            item.addresses.map((address) => propertyAddressKey(part.name, address))
          )
      );
    const visit = (node: OoxmlNode, path: readonly number[]) => {
      paths.set(keyOf(part.name, node.id), path);
      if (node.kind === 'textValue') return;
      const revision = revisions.get(node.id);
      const address = revision ? groupAddress(node) : null,
        inertHistory =
          revision &&
          !revision.refused &&
          address &&
          ['tblGridChange', 'tblPrChange'].includes(node.localName) &&
          !active.has(propertyAddressKey(part.name, address));
      if (revision && !inertHistory) {
        const location = locations.get(node.id),
          site: Site = {
            key: keyOf(part.name, node.id),
            partName: part.name,
            node,
            address,
            path,
            range: location
              ? {
                  partName: part.name,
                  start: { paragraphId: location.paragraphId, offset: location.start },
                  end: { paragraphId: location.paragraphId, offset: location.end },
                }
              : null,
            refused:
              revision.refused ||
              !address ||
              (inherited.get(keyOf(part.name, node.id))?.readOnly ?? false),
            kind: revision.propertyChange
              ? 'format'
              : revision.paragraphMark
                ? 'paragraphMark'
                : node.kind,
            marker: reviewGroupMarker(node) ?? inherited.get(keyOf(part.name, node.id))?.id,
          };
        sites.push(site);
        byKey.set(site.key, site);
        if (address) {
          const key = propertyAddressKey(part.name, address),
            values = byAddress.get(key) ?? [];
          values.push(site);
          byAddress.set(key, values);
        }
      }
      node.children.forEach((child, index) => visit(child, [...path, index]));
    };
    visit(part.root, []);
  }
  const newBucket = (id?: string): Bucket => {
    const bucket: Bucket = { id, sites: [], properties: [], readOnly: false };
    buckets.push(bucket);
    return bucket;
  };
  const attach = (bucket: Bucket, site: Site) => {
    const existing = owner.get(site.key);
    if (existing && existing !== bucket)
      throw new Error('A native revision site has conflicting group ownership.');
    if (site.marker && bucket.id && site.marker !== bucket.id)
      throw new Error('An application proposal conflicts with persisted native group ownership.');
    if (!existing) {
      bucket.sites.push(site);
      bucket.readOnly ||= site.refused;
      owner.set(site.key, bucket);
    }
  };
  const addressedSites = (keys: ReadonlySet<string>) => [
    ...new Set([...keys].flatMap((key) => byAddress.get(key) ?? [])),
  ];
  const store = readPropertyReviewStore(pkg),
    index = buildBlockIndex(pkg);
  if (store.status === 'invalid') throw new Error(store.reason);
  const decisions = readPayloadGroupDecisions(pkg);
  if (decisions.status === 'invalid') throw new Error(decisions.reason);
  const presenceOwners = new Map<string, number>(),
    propertyOwners = new Set<string>();
  if (store.status === 'ready')
    for (const entry of store.records.values()) {
      const record = entry.decoded.status === 'supported' ? entry.decoded.value : null;
      const presence = record ? null : decodeBlockPresenceRecord(entry.raw);
      if (record) {
        for (const block of record.blocks)
          if (
            !decisions.decisions.has(payloadGroupDecisionKey(entry.id, block.blockId)) ||
            decisions.decisions.get(payloadGroupDecisionKey(entry.id, block.blockId))?.status ===
              'conflict'
          )
            propertyOwners.add(block.blockId);
      } else if (presence?.status === 'supported') {
        for (const block of presence.value.blocks)
          if (
            !decisions.decisions.has(payloadGroupDecisionKey(entry.id, block.blockId)) ||
            decisions.decisions.get(payloadGroupDecisionKey(entry.id, block.blockId))?.status ===
              'conflict'
          )
            for (const member of block.members)
              presenceOwners.set(
                member.payload.id,
                (presenceOwners.get(member.payload.id) ?? 0) + 1
              );
      }
    }
  if (store.status === 'ready')
    for (const entry of store.records.values()) {
      if (entry.decoded.status !== 'supported') {
        const presence = decodeBlockPresenceRecord(entry.raw);
        if (presence.status !== 'supported')
          throw new Error('Unsupported application proposal version in review groups.');
        for (const block of presence.value.blocks) {
          const disposition = decisions.decisions.get(
            payloadGroupDecisionKey(entry.id, block.blockId)
          );
          if (disposition && disposition.status !== 'conflict') continue;
          const bucket = newBucket(payloadGroupId(entry.id, block.blockId));
          bucket.properties.push({ proposalId: entry.id, blockId: block.blockId });
          const native = index.byId.get(block.blockId),
            path = native && paths.get(keyOf(native.partName, native.controlId));
          bucket.readOnly =
            !!disposition ||
            block.members.some(
              (member) =>
                propertyOwners.has(member.payload.id) ||
                (presenceOwners.get(member.payload.id) ?? 0) > 1
            ) ||
            !bindBlockPresenceRecord(pkg, { ...presence.value, blocks: [block] }).success;
          if (!native || !path) {
            bucket.readOnly = true;
            continue;
          }
          const fallback = atomicFallbackParagraph(
            pkg.parts.get(native.partName)!,
            native.controlId
          );
          bucket.anchor = {
            partName: native.partName,
            nodeId: native.controlId,
            path,
            paragraphId: fallback?.id ?? null,
            offset: 0,
          };
          marked.set(bucket.id!, bucket);
        }
        continue;
      }
      const record = entry.decoded.value;
      for (const block of record.blocks) {
        const disposition = decisions.decisions.get(
          payloadGroupDecisionKey(record.id, block.blockId)
        );
        if (disposition && disposition.status !== 'conflict') continue;
        const bucket = newBucket(payloadGroupId(record.id, block.blockId)),
          native = index.byId.get(block.blockId),
          path = native && paths.get(keyOf(native.partName, native.controlId));
        bucket.properties.push({ proposalId: record.id, blockId: block.blockId });
        if (disposition || presenceOwners.has(block.blockId)) bucket.readOnly = true;
        marked.set(bucket.id!, bucket);
        if (!native || !path || native.payload.type !== block.type) {
          bucket.readOnly = true;
          continue;
        }
        const fallback =
          propertyFallbackParagraph(pkg.parts.get(native.partName)!, native.controlId) ??
          atomicFallbackParagraph(pkg.parts.get(native.partName)!, native.controlId);
        bucket.anchor = {
          partName: native.partName,
          nodeId: native.controlId,
          path,
          paragraphId: fallback?.id ?? null,
          offset: 0,
        };
        const addresses = new Set(
          record.nativeRevisions.map((address) => propertyAddressKey(address.partName, address))
        );
        applicationScopes.set(bucket, {
          partName: native.partName,
          path,
          paragraphId: fallback?.id ?? null,
          addresses,
        });
        for (const site of addressedSites(addresses))
          if (
            site.partName === native.partName &&
            within(site.path, path) &&
            fallback &&
            site.range?.start.paragraphId === fallback.id
          )
            attach(bucket, site);
        const labels = record.extension?.fallbackLabels,
          label = Array.isArray(labels)
            ? labels.find(
                (value) =>
                  value &&
                  typeof value === 'object' &&
                  !Array.isArray(value) &&
                  value.blockId === block.blockId
              )
            : null,
          needsNative =
            label && typeof label === 'object' && !Array.isArray(label)
              ? label.textBefore !== label.textAfter
              : record.nativeRevisions.length > 0;
        if (needsNative && !bucket.sites.length) bucket.readOnly = true;
      }
    }
  const groupStore = readReviewGroupStore(pkg);
  if (groupStore.status === 'invalid') throw new Error(groupStore.reason);
  if (groupStore.status === 'ready')
    for (const entry of groupStore.records.values()) {
      if (entry.decoded.status === 'payload-decision') continue;
      if (entry.decoded.status !== 'supported')
        throw new Error('Unsupported persisted review group version.');
      const record = entry.decoded.value,
        keys = new Set(
          record.members.map((member) => propertyAddressKey(member.partName, member.revision))
        ),
        members = addressedSites(keys);
      // Decisions remove native members. Keep their record for native decision undo.
      if (!members.length) continue;
      const id = `g1:stored:${record.id}`,
        bucket = newBucket(id);
      marked.set(id, bucket);
      bucket.readOnly = [...keys].some((key) => !byAddress.has(key));
      for (const site of members) attach(bucket, site);
    }
  for (const site of sites) {
    if (owner.has(site.key) || !site.marker) continue;
    const bucket = marked.get(site.marker) ?? newBucket(site.marker);
    marked.set(site.marker, bucket);
    if (site.marker.startsWith('g1:payload:')) {
      const scope = applicationScopes.get(bucket);
      if (
        !scope ||
        site.partName !== scope.partName ||
        !within(site.path, scope.path) ||
        site.range?.start.paragraphId !== scope.paragraphId ||
        !site.address ||
        !scope.addresses.has(propertyAddressKey(site.partName, site.address))
      )
        bucket.readOnly = true;
    }
    attach(bucket, site);
  }
  for (const receipt of receipts) {
    const addresses = new Set(
        receipt.addresses.map((address) => propertyAddressKey(receipt.partName, address))
      ),
      members = addressedSites(addresses);
    const existing = new Set(
      members.map((site) => owner.get(site.key)).filter((bucket) => bucket !== undefined)
    );
    if (existing.size > 1)
      throw new Error('A native receipt crosses independent persisted review groups.');
    const bucket = [...existing][0] ?? newBucket();
    const ownedAddresses = new Set(
      bucket.sites.flatMap((site) =>
        site.address ? [propertyAddressKey(site.partName, site.address)] : []
      )
    );
    if (
      bucket.id &&
      members.some(
        (site) =>
          !owner.has(site.key) &&
          (!site.address || !ownedAddresses.has(propertyAddressKey(site.partName, site.address)))
      )
    )
      throw new Error('A native receipt cannot extend an independent persisted review group.');
    for (const site of members) attach(bucket, site);
  }
  // Core retains same-kind fragments and structural decisions. Persisted groups always take precedence.
  for (const part of parts)
    for (const item of itemsByPart.get(part.name)!) {
      const addresses = new Set(
          item.addresses.map((address) => propertyAddressKey(part.name, address))
        ),
        members = addressedSites(addresses).filter(
          (site) =>
            !owner.has(site.key) &&
            (item.revisionKind === 'format'
              ? site.kind === 'format'
              : item.revisionKind === 'insert'
                ? ['revisionInsert', 'paragraphMark'].includes(site.kind)
                : item.revisionKind === 'delete'
                  ? ['revisionDelete', 'paragraphMark'].includes(site.kind)
                  : item.revisionKind === 'paragraphMark'
                    ? site.kind === 'paragraphMark'
                    : true) &&
            (!site.range ||
              item.ranges.some(
                (range) =>
                  range.start.paragraphId === site.range!.start.paragraphId &&
                  range.start.offset <= site.range!.start.offset &&
                  range.end.paragraphId === site.range!.end.paragraphId &&
                  range.end.offset >= site.range!.end.offset
              ))
        );
      if (!members.length) continue;
      const bucket = newBucket();
      bucket.readOnly ||= item.readOnly;
      for (const site of members) attach(bucket, site);
    }
  for (const site of sites) if (!owner.has(site.key)) attach(newBucket(), site);
  // Consecutive native insertion/deletion wrappers form one replacement at every nesting depth.
  const merge = (left: Bucket, right: Bucket) => {
    if (left === right || left.id || right.id) return;
    for (const site of right.sites) {
      owner.set(site.key, left);
      left.sites.push(site);
    }
    left.readOnly ||= right.readOnly;
    right.sites = [];
  };
  for (const part of parts) {
    const visit = (node: OoxmlNode) => {
      if (node.kind === 'textValue') return;
      let previous: Site | undefined;
      for (const child of node.children) {
        const site = byKey.get(keyOf(part.name, child.id));
        if (site && ['revisionInsert', 'revisionDelete'].includes(child.kind)) {
          if (
            previous?.address &&
            site.address &&
            previous.address.author === site.address.author &&
            previous.address.date === site.address.date
          )
            merge(owner.get(previous.key)!, owner.get(site.key)!);
          previous = site;
        } else previous = undefined;
        visit(child);
      }
    };
    visit(part.root);
  }
  const compareAnchor = (a: ReviewGroupAnchor, b: ReviewGroupAnchor) =>
    partOrder.get(a.partName)! - partOrder.get(b.partName)! || pathCompare(a.path, b.path);
  const groups = buckets
    .filter((bucket) => bucket.sites.length || bucket.properties.length)
    .map((bucket): NativeReviewGroup => {
      bucket.sites.sort((a, b) => compareAnchor(anchorOf(a), anchorOf(b)));
      const first = bucket.sites[0],
        anchor = bucket.anchor ?? (first ? anchorOf(first) : null),
        id =
          bucket.id ??
          `g1:native:${bucket.sites.map((site) => JSON.stringify([site.partName, site.address, site.kind])).sort()[0]!}`,
        members = new Map<
          string,
          { partName: string; address: RevisionAddress; siteNodeIds: string[] }
        >();
      for (const site of bucket.sites) {
        if (!site.address) continue;
        const key = propertyAddressKey(site.partName, site.address),
          member = members.get(key) ?? {
            partName: site.partName,
            address: { ...site.address },
            siteNodeIds: [],
          };
        member.siteNodeIds.push(site.node.id);
        members.set(key, member);
      }
      return Object.freeze({
        id,
        native: Object.freeze(
          [...members.values()].map((member) =>
            Object.freeze({
              ...member,
              address: Object.freeze(member.address),
              siteNodeIds: Object.freeze(member.siteNodeIds),
            })
          )
        ),
        properties: Object.freeze(bucket.properties.map((property) => Object.freeze(property))),
        anchor: anchor ? Object.freeze({ ...anchor, path: Object.freeze([...anchor.path]) }) : null,
        readOnly: bucket.readOnly,
      });
    })
    .sort((a, b) =>
      a.anchor && b.anchor
        ? compareAnchor(a.anchor, b.anchor) || a.id.localeCompare(b.id)
        : a.anchor
          ? -1
          : b.anchor
            ? 1
            : a.id.localeCompare(b.id)
    );
  const counts = new Map<string, number>();
  for (const group of groups) counts.set(group.id, (counts.get(group.id) ?? 0) + 1);
  const ordinal = new Map<string, number>();
  // Imported files can reuse identical native addresses for independent decisions. Keep each
  // visible, but require explicit ownership reconciliation before decisions on those groups.
  const identified = groups.map((group) => {
    if (counts.get(group.id) === 1) return group;
    if (buckets.some((bucket) => bucket.id === group.id))
      throw new Error('Duplicate persisted review group identities.');
    const index = ordinal.get(group.id) ?? 0;
    ordinal.set(group.id, index + 1);
    return Object.freeze({ ...group, id: `${group.id}:imported:${index}`, readOnly: true });
  });
  return { groups: Object.freeze(identified), sites };
}

/**
 * Persist new pairing in the tracked edit transaction. Existing imported native XML stays exact.
 * Prior group identities constrain new grouping, so a new neighbor cannot merge earlier decisions.
 */
export function persistReviewGroups(
  before: OoxmlPackage,
  after: OoxmlPackage,
  receipts: readonly ReviewGroupReceipt[] = [],
  allowInitialization = false
): OoxmlPackage {
  const observed = derive(before),
    inherited = new Map<string, { id: string; readOnly: boolean }>();
  for (const group of observed.groups)
    for (const member of group.native)
      for (const nodeId of member.siteNodeIds)
        inherited.set(keyOf(member.partName, nodeId), { id: group.id, readOnly: group.readOnly });
  const current = derive(after, receipts, inherited),
    priorIds = new Set(observed.groups.map((group) => group.id)),
    liveAddresses = new Set(
      observed.sites.flatMap((site) =>
        site.address ? [propertyAddressKey(site.partName, site.address)] : []
      )
    ),
    retiredAddresses = new Set<string>(),
    groupStore = readReviewGroupStore(before);
  if (groupStore.status === 'ready')
    for (const entry of groupStore.records.values()) {
      if (entry.decoded.status !== 'supported') continue; // derive already refuses future records.
      for (const member of entry.decoded.value.members) {
        const key = propertyAddressKey(member.partName, member.revision);
        if (!liveAddresses.has(key)) retiredAddresses.add(key);
      }
    }
  if (
    current.sites.some(
      (site) =>
        !inherited.has(site.key) &&
        site.address &&
        retiredAddresses.has(propertyAddressKey(site.partName, site.address))
    )
  )
    throw new Error('A new native revision reuses an address retained by review history.');
  const created = current.groups.filter(
    (group) =>
      !priorIds.has(group.id) &&
      !group.properties.length &&
      group.native.some((member) =>
        member.siteNodeIds.some((nodeId) => !inherited.has(keyOf(member.partName, nodeId)))
      )
  );
  if (!created.length) return after;
  let result = after;
  if (allowInitialization) {
    // Only a solo host can own part creation in its own history. Shared hosts prepare before joining.
    const initialized = withInitializedReviewGroupStore(result);
    if (!initialized.ok) throw new Error(initialized.reason);
    result = initialized.pkg;
  }
  for (const group of created) {
    if (group.readOnly || !group.native.length)
      throw new Error(`A new review group needs exact native membership (${group.id}).`);
    const written = withReviewGroupRecord(result, {
      schemaVersion: 1,
      id: `rg_${crypto.randomUUID().replaceAll('-', '')}`,
      members: group.native.map((member) => ({
        partName: member.partName,
        revision: member.address,
      })),
    });
    if (!written.ok) throw new Error(written.reason);
    result = written.pkg;
  }
  return result;
}
