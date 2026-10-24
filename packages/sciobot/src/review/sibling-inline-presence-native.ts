import {
  collectRevisionSites,
  findNode,
  parentNodeOf,
  paragraphTextOf,
  revisionItemsOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { nativePropertyAddress, paragraphOwner, propertyAddressKey } from './property-fallback';
import type { BoundInlinePresenceNative, InlinePresenceNative } from './inline-presence-native';

/** Prove one complete native sibling gap, including all rows of each newly born table. */
export function bindSiblingInlinePresenceNative(
  pkg: OoxmlPackage,
  blockId: string,
  partName: string,
  controlId: string,
  expected: InlinePresenceNative
): BoundInlinePresenceNative {
  const part = pkg.parts.get(partName);
  if (
    !part ||
    !expected.siblings ||
    expected.cell ||
    expected.kind !== 'ins' ||
    expected.blockId !== blockId ||
    expected.partName !== partName ||
    !expected.gap?.length ||
    expected.gap.length > 1024
  )
    throw Error('Sibling atom ownership needs an exact bounded structural receipt.');
  const addresses = new Set(expected.gap.map((a) => propertyAddressKey(partName, a)));
  const ownKey = propertyAddressKey(partName, expected.revision);
  if (
    addresses.size !== expected.gap.length ||
    !addresses.has(ownKey) ||
    expected.gap.some(
      (a) => a.author !== expected.revision.author || a.date !== expected.revision.date
    )
  )
    throw Error('Sibling atom ownership has duplicate, foreign, or missing native addresses.');
  let at = parentNodeOf(part, controlId),
    wrapper: OoxmlNode | null = null;
  while (at && at.kind !== 'paragraph') {
    const address = nativePropertyAddress(at);
    if (at.kind === 'revisionInsert' && address && propertyAddressKey(partName, address) === ownKey)
      wrapper = at;
    at = parentNodeOf(part, at.id);
  }
  if (!at || !wrapper) throw Error('The sibling atom left its paragraph insertion.');
  const paragraph = at;
  // A top-level new table and its following caret paragraph share the native sibling receipt.
  let root: OoxmlNode = paragraph,
    container = parentNodeOf(part, root.id);
  while (container && !['body', 'tc', 'sdtContent'].includes(container.localName)) {
    root = container;
    container = parentNodeOf(part, root.id);
  }
  // A paragraph in a cell belongs to the outer table gap, not its cell's inner story.
  if (container?.kind === 'tableCell') {
    const row = parentNodeOf(part, container.id),
      table = row && parentNodeOf(part, row.id);
    if (row?.kind !== 'tableRow' || table?.kind !== 'table')
      throw Error('Invalid born table ancestry.');
    root = table;
    container = parentNodeOf(part, root.id);
  }
  if (
    !container ||
    container.namespaceUri !== WML_NAMESPACE_URI ||
    !['body', 'tc', 'sdtContent'].includes(container.localName)
  )
    throw Error('Sibling atom ownership has no supported native story.');
  const ownerOf = (node: OoxmlNode): OoxmlNode | null => {
    let owner: OoxmlNode | null = node;
    for (let depth = 0; owner && depth < 128; depth++) {
      const parent = parentNodeOf(part, owner.id);
      if (parent?.id === container!.id) return owner;
      owner = parent;
    }
    return null;
  };
  const matched = new Set<string>(),
    roots = new Map<string, OoxmlNode>();
  const allSites = collectRevisionSites(part);
  const sites = allSites.filter((site) => {
    const address = nativePropertyAddress(site.node);
    if (!address || !addresses.has(propertyAddressKey(partName, address))) return false;
    const owner = ownerOf(site.node);
    if (
      site.refused ||
      !['ins', 'del', 'cellIns'].includes(site.node.localName) ||
      !owner ||
      !['paragraph', 'table'].includes(owner.kind)
    )
      throw Error('Sibling atom membership escapes its native gap.');
    roots.set(owner.id, owner);
    matched.add(propertyAddressKey(partName, address));
    return true;
  });
  if (
    matched.size !== addresses.size ||
    !sites.some((s) => s.node.id === wrapper!.id) ||
    !sites.some(
      (s) =>
        s.paragraphMark &&
        s.node.localName === 'ins' &&
        paragraphOwner(part, s.node) === paragraph.id &&
        propertyAddressKey(partName, nativePropertyAddress(s.node)!) === ownKey
    )
  )
    throw Error('Sibling atom ownership lost its native paragraph birth.');
  for (const item of revisionItemsOf(part))
    if (
      item.addresses.some((a) => addresses.has(propertyAddressKey(partName, a))) &&
      item.addresses.some((a) => !addresses.has(propertyAddressKey(partName, a)))
    )
      throw Error('Sibling atom ownership omits a coupled native dependency.');
  const positions = [...roots.keys()].map((id) =>
    container!.children.findIndex((n) => n.id === id)
  );
  const start = Math.min(...positions),
    end = Math.max(...positions);
  if (start < 0 || container.children.slice(start, end + 1).some((n) => !roots.has(n.id)))
    throw Error('Sibling atom ownership crosses an unrelated native root.');
  const ownedMarks = new Set(
    sites.filter((s) => s.paragraphMark).map((s) => paragraphOwner(part, s.node))
  );
  if (![...roots.values()].some((owner) => owner.kind === 'table'))
    throw Error('Sibling table ownership has no whole table birth.');
  for (const site of allSites) {
    const address = nativePropertyAddress(site.node);
    if (
      address?.author !== expected.revision.author ||
      address.date !== expected.revision.date ||
      addresses.has(propertyAddressKey(partName, address))
    )
      continue;
    const owner = ownerOf(site.node);
    if (owner && roots.has(owner.id))
      throw Error('Sibling atom ownership omits a native root dependency.');
  }
  for (const owner of roots.values()) {
    if (owner.kind === 'paragraph') {
      if (!ownedMarks.has(owner.id))
        throw Error('Sibling atom ownership claims an unowned paragraph.');
    } else if (owner.kind === 'table') {
      for (const row of owner.children.filter((n) => n.kind === 'tableRow')) {
        const birth = sites.some(
          (s) =>
            s.node.localName === 'ins' &&
            parentNodeOf(part, s.node.id)?.localName === 'trPr' &&
            parentNodeOf(part, parentNodeOf(part, s.node.id)!.id)?.id === row.id
        );
        if (!birth) throw Error('Sibling atom ownership claims a table with an unowned row.');
      }
      const visit = (node: OoxmlNode) => {
        if (node.kind === 'textValue') return;
        if (node.kind === 'paragraph' && paragraphTextOf(part, node.id) && !ownedMarks.has(node.id))
          throw Error('Sibling atom ownership claims unowned table content.');
        node.children.forEach(visit);
      };
      visit(owner);
    }
  }
  // Roots are derived from persistent revision membership, never serialized canonical IDs.
  if (!roots.has(root.id) || !findNode(part, controlId))
    throw Error('Sibling atom ownership lost its root.');
  return { ...expected, siteNodeId: wrapper.id, siteNodeIds: sites.map((s) => s.node.id) };
}
