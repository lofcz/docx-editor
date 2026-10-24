import { collectRevisionSites, parentNodeOf, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { nativePropertyAddress, paragraphOwner, propertyAddressKey } from './property-fallback';
import type { BoundInlinePresenceNative, InlinePresenceNative } from './inline-presence-native';

/** Bind the whole structural receipt, while proving this control belongs to its inserted paragraph. */
export function bindBornInlinePresenceNative(
  pkg: OoxmlPackage,
  blockId: string,
  partName: string,
  controlId: string,
  expected: InlinePresenceNative
): BoundInlinePresenceNative {
  if (expected.kind !== 'ins' || !expected.gap?.length || expected.gap.length > 1024)
    throw Error('Born inline ownership needs an insertion and a bounded structural receipt.');
  if (expected.blockId !== blockId || expected.partName !== partName)
    throw Error('Born inline ownership changed its control or part.');
  const part = pkg.parts.get(partName);
  if (!part) throw Error('Born inline ownership has no native part.');
  const key = propertyAddressKey(partName, expected.revision);
  let parent = parentNodeOf(part, controlId),
    wrapperId: string | null = null;
  while (parent && parent.kind !== 'paragraph') {
    const address = nativePropertyAddress(parent);
    if (
      parent.kind === 'revisionInsert' &&
      address &&
      propertyAddressKey(partName, address) === key
    )
      wrapperId = parent.id;
    parent = parentNodeOf(part, parent.id);
  }
  if (!parent || !wrapperId) throw Error('The born control left its paragraph insertion.');
  const paragraph = parent;
  const addresses = new Set(expected.gap.map((address) => propertyAddressKey(partName, address)));
  if (addresses.size !== expected.gap.length || !addresses.has(key))
    throw Error('Born inline ownership has duplicate or missing native addresses.');
  if (
    expected.gap.some(
      (address) =>
        address.author !== expected.revision.author || address.date !== expected.revision.date
    )
  )
    throw Error('Paragraph birth membership has foreign attribution.');
  const matched = new Set<string>();
  const sites = collectRevisionSites(part).filter((site) => {
    const address = nativePropertyAddress(site.node);
    if (!address) return false;
    const member = propertyAddressKey(partName, address);
    if (!addresses.has(member)) return false;
    if (site.refused || !['ins', 'del'].includes(site.node.localName))
      throw Error('Born inline ownership has unsupported structural membership.');
    matched.add(member);
    return true;
  });
  const marks = collectRevisionSites({ ...part, root: paragraph }).filter((site) => {
    const address = nativePropertyAddress(site.node);
    return (
      site.paragraphMark &&
      site.node.localName === 'ins' &&
      address &&
      propertyAddressKey(partName, address) === key
    );
  });
  if (
    matched.size !== addresses.size ||
    marks.length !== 1 ||
    !sites.some((site) => site.node.id === wrapperId)
  )
    throw Error('Born inline ownership lost its paragraph mark or native members.');
  const container = parentNodeOf(part, paragraph.id);
  if (!container) throw Error('Paragraph birth has no native container.');
  const owners = new Set<string>();
  const markedOwners = new Set<string>();
  for (const site of sites) {
    const owner = paragraphOwner(part, site.node);
    if (!owner || parentNodeOf(part, owner)?.id !== container.id)
      throw Error('Paragraph birth membership crosses native containers.');
    owners.add(owner);
    if (site.paragraphMark) markedOwners.add(owner);
  }
  if ([...owners].some((owner) => !markedOwners.has(owner)))
    throw Error('Paragraph birth membership claims unowned paragraph text.');
  const positions = [...owners].map((owner) =>
    container.children.findIndex((node) => node.id === owner)
  );
  const from = positions.reduce((minimum, position) => Math.min(minimum, position), Infinity),
    to = positions.reduce((maximum, position) => Math.max(maximum, position), -1);
  if (
    container.children
      .slice(from, to + 1)
      .some((node) => node.kind === 'paragraph' && !owners.has(node.id))
  )
    throw Error('Paragraph birth membership crosses an unrelated paragraph.');
  return { ...expected, siteNodeId: wrapperId, siteNodeIds: sites.map((site) => site.node.id) };
}
