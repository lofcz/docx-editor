import {
  collectRevisionSites,
  revisionItemsOf,
  parentNodeOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import { nativePropertyAddress, propertyAddressKey, paragraphOwner } from './property-fallback';
import type { InlinePresenceNative, BoundInlinePresenceNative } from './inline-presence-native';

export function cellRevisionAddress(node: OoxmlNode): RevisionAddress | null {
  if (node.kind === 'textValue') return null;
  if (node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'tblGridChange') {
    const id = node.attributes.find(
      (a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'id'
    )?.value;
    return id === undefined ? null : { id, author: '' };
  }
  return nativePropertyAddress(node);
}
/** The whole table receipt owns a cell atom; paragraph ownership alone cannot decide topology. */
export function bindCellInlinePresenceNative(
  pkg: OoxmlPackage,
  blockId: string,
  partName: string,
  controlId: string,
  expected: InlinePresenceNative
): BoundInlinePresenceNative {
  const part = pkg.parts.get(partName),
    cellOwner = expected.cell;
  if (
    !part ||
    !cellOwner ||
    expected.blockId !== blockId ||
    expected.partName !== partName ||
    expected.kind !== 'ins' ||
    !expected.gap?.length ||
    expected.gap.length > 1024
  )
    throw Error('Cell atom ownership needs an exact bounded table receipt.');
  let cell: OoxmlNode | null = parentNodeOf(part, controlId);
  while (cell && cell.kind !== 'tableCell') cell = parentNodeOf(part, cell.id);
  const row = cell && parentNodeOf(part, cell.id);
  const table = row && parentNodeOf(part, row.id);
  if (cell?.kind !== 'tableCell' || row?.kind !== 'tableRow' || table?.kind !== 'table')
    throw Error('Cell atom ownership lost its native owners.');
  let at = parentNodeOf(part, controlId),
    wrapperId: string | null = null,
    paragraphId: string | null = null;
  const key = propertyAddressKey(partName, expected.revision);
  while (at && at.id !== cell.id) {
    if (at.kind === 'revisionInsert') {
      const address = nativePropertyAddress(at);
      if (address && propertyAddressKey(partName, address) === key) wrapperId = at.id;
    }
    if (at.kind === 'paragraph') paragraphId = at.id;
    at = parentNodeOf(part, at.id);
  }
  if (!at || !wrapperId || !paragraphId) throw Error('The cell atom left its paragraph insertion.');
  const addresses = new Set(expected.gap.map((a) => propertyAddressKey(partName, a)));
  if (addresses.size !== expected.gap.length || !addresses.has(key))
    throw Error('Cell ownership has duplicate or missing addresses.');
  const matched = new Set<string>(),
    sites = collectRevisionSites(part).filter((site) => {
      const address = cellRevisionAddress(site.node);
      if (!address || !addresses.has(propertyAddressKey(partName, address))) return false;
      if (
        site.refused ||
        (address.author !== expected.revision.author &&
          !(site.node.localName === 'tblGridChange' && address.author === '')) ||
        (address.author && address.date !== expected.revision.date)
      )
        throw Error('Cell ownership has foreign or unsupported membership.');
      let owner: OoxmlNode | null = site.node;
      while (owner && owner.id !== table.id) owner = parentNodeOf(part, owner.id);
      if (!owner) throw Error('Cell ownership escapes its native table.');
      matched.add(propertyAddressKey(partName, address));
      return true;
    });
  const topologyKey = propertyAddressKey(partName, cellOwner.revision);
  for (const item of revisionItemsOf(part)) {
    if (!item.addresses.some((address) => propertyAddressKey(partName, address) === topologyKey))
      continue;
    if (item.addresses.some((address) => !addresses.has(propertyAddressKey(partName, address))))
      throw Error('Cell ownership omits a native topology dependency.');
  }
  const scope = cellOwner.kind === 'row' ? row : cell;
  for (const site of collectRevisionSites({ ...part, root: scope })) {
    const address = cellRevisionAddress(site.node);
    if (
      address &&
      address.author === expected.revision.author &&
      address.date === expected.revision.date &&
      !addresses.has(propertyAddressKey(partName, address))
    )
      throw Error('Cell ownership omits a native birth dependency.');
  }
  if (
    cellOwner.kind === 'cell' &&
    sites.some((site) => {
      const address = cellRevisionAddress(site.node);
      const parent = parentNodeOf(part, site.node.id);
      return (
        site.node.localName === 'ins' &&
        parent?.localName === 'trPr' &&
        parentNodeOf(part, parent.id)?.id === row.id &&
        address &&
        propertyAddressKey(partName, address) === topologyKey
      );
    })
  )
    throw Error('A born row requires its row topology proof.');
  const birth = sites.some((site) => {
    const address = cellRevisionAddress(site.node);
    if (!address || propertyAddressKey(partName, address) !== topologyKey) return false;
    if (cellOwner.kind === 'cell' && site.node.localName === 'cellIns')
      return parentNodeOf(part, parentNodeOf(part, site.node.id)?.id ?? '')?.id === cell.id;
    if (cellOwner.kind !== 'row' || site.node.localName !== 'ins') return false;
    const parent = parentNodeOf(part, site.node.id);
    return parent?.localName === 'trPr' && parentNodeOf(part, parent.id)?.id === row.id;
  });
  const mark = sites.some(
    (site) =>
      site.paragraphMark &&
      site.node.localName === 'ins' &&
      cellRevisionAddress(site.node) &&
      propertyAddressKey(partName, cellRevisionAddress(site.node)!) === key &&
      paragraphOwner(part, site.node) === paragraphId
  );
  if (
    !birth ||
    !mark ||
    matched.size !== addresses.size ||
    !sites.some((s) => s.node.id === wrapperId)
  )
    throw Error('Cell atom ownership lost its topology or paragraph birth.');
  return { ...expected, siteNodeId: wrapperId, siteNodeIds: sites.map((site) => site.node.id) };
}
