import {
  collectRevisionSites,
  customXmlNodes,
  withCustomXmlNode,
  writeOoxmlPackage,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlPackage,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import { readReviewGroupStore } from './group-record-store';
import {
  encodeReviewGroupRecord,
  reviewGroupMemberKey,
  type StoredReviewGroupMember,
} from './group-record-schema';
import { nativePropertyAddress } from './property-fallback';
import { activeStorage } from '../registry/active';
import { readPropertyReviewStore } from './property-record-store';
import { encodePropertyReviewRecord } from './property-record-schema';
import { decodeBlockPresenceRecord, encodeBlockPresenceRecord } from './block-presence-schema';
import { inlinePresenceExtension, readInlinePresenceExtension } from './inline-presence-native';

function addressOf(node: OoxmlElement): RevisionAddress | null {
  const address = nativePropertyAddress(node);
  if (address || node.localName !== 'tblGridChange') return address;
  const id = node.attributes.find(
    (a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'id'
  )?.value;
  return id === undefined ? null : { id, author: '' };
}

/** Rebind persisted native group addresses in an export snapshot. Live records stay immutable. */
export function exportReviewGroupReferences(
  source: OoxmlPackage,
  exported: OoxmlPackage
): OoxmlPackage {
  if (source === exported) return exported;
  const store = readReviewGroupStore(source);
  const remapped = new Map<string, readonly StoredReviewGroupMember[]>();
  const scanned = new Set<string>();
  const membersOf = (member: StoredReviewGroupMember): readonly StoredReviewGroupMember[] => {
    const partName = member.partName;
    if (!scanned.has(partName)) {
      scanned.add(partName);
      const before = source.parts.get(partName),
        after = exported.parts.get(partName);
      if (before && after && before !== after) {
        const targets = new Map(
          collectRevisionSites(after).map((site) => [site.node.id, site.node])
        );
        const grouped = new Map<string, Map<string, StoredReviewGroupMember>>();
        for (const site of collectRevisionSites(before)) {
          const old = addressOf(site.node),
            target = targets.get(site.node.id);
          const revision = target ? addressOf(target) : null;
          if (!old || !revision) continue;
          const key = reviewGroupMemberKey({ partName, revision: old });
          let rows = grouped.get(key);
          if (!rows) grouped.set(key, (rows = new Map()));
          const next = { partName, revision };
          rows.set(reviewGroupMemberKey(next), next);
        }
        for (const [key, rows] of grouped) remapped.set(key, [...rows.values()]);
      }
    }
    return remapped.get(reviewGroupMemberKey(member)) ?? [member];
  };
  let result = exported;
  if (store.status === 'ready')
    for (const record of store.retainedRecords.values()) {
      if (record.decoded.status !== 'supported') continue;
      const value = record.decoded.value;
      const members = new Map<string, StoredReviewGroupMember>();
      for (const member of value.members)
        for (const next of membersOf(member)) members.set(reviewGroupMemberKey(next), next);
      const data = encodeReviewGroupRecord({ ...value, members: [...members.values()] });
      if (data !== record.raw)
        result = withCustomXmlNode(result, store.part.partName, {
          id: record.id,
          label: store.removedIds.has(record.id) ? 'removed' : '',
          data,
        });
    }
  const properties = readPropertyReviewStore(source);
  if (properties.status !== 'ready') return result;
  const labels = new Map(
    customXmlNodes(source, properties.part.partName).map((node) => [node.id, node.label])
  );
  for (const record of properties.retainedRecords.values()) {
    let data: string;
    if (record.decoded.status === 'supported') {
      const value = record.decoded.value;
      const nativeRevisions = value.nativeRevisions.flatMap(({ partName, ...revision }) =>
        membersOf({ partName, revision }).map((next) => ({ partName, ...next.revision }))
      );
      data = encodePropertyReviewRecord({ ...value, nativeRevisions });
    } else {
      const presence = decodeBlockPresenceRecord(record.raw);
      if (presence.status !== 'supported') continue;
      const owners = readInlinePresenceExtension(presence.value);
      if (!owners) continue;
      const mapped = [...owners.values()].map((owner) => {
        const addresses = (revision: RevisionAddress) =>
          membersOf({ partName: owner.partName, revision }).map((next) => next.revision);
        const single = (revision: RevisionAddress) => {
          const values = addresses(revision);
          if (values.length !== 1)
            throw Error('Export cannot preserve ambiguous native inline ownership.');
          return values[0]!;
        };
        return {
          ...owner,
          revision: single(owner.revision),
          ...(owner.gap ? { gap: owner.gap.flatMap(addresses) } : {}),
          ...(owner.cell ? { cell: { ...owner.cell, revision: single(owner.cell.revision) } } : {}),
        };
      });
      data = encodeBlockPresenceRecord({
        ...presence.value,
        extension: { ...presence.value.extension, ...inlinePresenceExtension(mapped) },
      });
    }
    if (data !== record.raw)
      result = withCustomXmlNode(result, properties.part.partName, {
        id: activeStorage().reviewNodePrefix + record.id,
        label: labels.get(activeStorage().reviewNodePrefix + record.id) ?? '',
        data,
      });
  }
  return result;
}

/** Export application group references together with native Word revision IDs. */
export function writeReviewGroupPackage(pkg: OoxmlPackage): Uint8Array {
  return writeOoxmlPackage(pkg, { exportRevisionReferences: exportReviewGroupReferences });
}
