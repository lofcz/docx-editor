import { bindCellInlinePresenceNative } from './cell-inline-presence-native';
import { bindBornInlinePresenceNative } from './born-inline-presence-native';
import {
  collectRevisionSites,
  parentNodeOf,
  type OoxmlNode,
  type OoxmlPackage,
  type RevisionAddress,
} from '@docx-editor.dev/core/store';
import { nativePropertyAddress, propertyAddressKey } from './property-fallback';
import type { BlockPresenceRecord } from './block-presence-schema';
import type { JsonObject } from '../data/json';

export interface InlinePresenceNative {
  readonly blockId: string;
  readonly partName: string;
  readonly kind: 'ins' | 'del';
  readonly revision: RevisionAddress;
  /** Version four: the complete paragraph birth/replacement receipt. */
  readonly gap?: readonly RevisionAddress[];
  /** Version five: exact native topology owner and its complete structural receipt. */
  readonly cell?: { readonly kind: 'row' | 'cell'; readonly revision: RevisionAddress };
}
export interface BoundInlinePresenceNative extends InlinePresenceNative {
  readonly siteNodeId: string;
  readonly siteNodeIds?: readonly string[];
}
const object = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const exact = (value: object, keys: readonly string[]) =>
  Object.keys(value).length === keys.length &&
  Object.keys(value).every((key) => keys.includes(key));

/** Native deleted text belongs to the outer decision, not to the immutable atom's source proof. */
export function inlinePresenceSource(node: OoxmlNode): OoxmlNode {
  if (node.kind === 'textValue') return node;
  const children = node.children.map(inlinePresenceSource);
  if (node.kind === 'deletedText')
    return { ...node, kind: 'text', localName: 't', children } as OoxmlNode;
  if (
    node.localName === 'delInstrText' &&
    node.namespaceUri === 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  )
    return { ...node, kind: 'instrText', localName: 'instrText', children } as OoxmlNode;
  return { ...node, children } as OoxmlNode;
}

/** Exactly one native wrapper owns exactly this one control. Never claim an adjacent text revision. */
export function bindInlinePresenceNative(
  pkg: OoxmlPackage,
  blockId: string,
  partName: string,
  controlId: string,
  kind: 'ins' | 'del',
  expected?: InlinePresenceNative
): BoundInlinePresenceNative {
  if (expected?.cell)
    return bindCellInlinePresenceNative(pkg, blockId, partName, controlId, expected);
  if (expected?.gap) {
    if (kind !== 'ins') throw Error('Paragraph birth ownership cannot own an inline deletion.');
    return bindBornInlinePresenceNative(pkg, blockId, partName, controlId, expected);
  }
  const part = pkg.parts.get(partName),
    parent = part && parentNodeOf(part, controlId);
  if (
    !part ||
    !parent ||
    parent.localName !== kind ||
    parent.kind !== (kind === 'ins' ? 'revisionInsert' : 'revisionDelete') ||
    parent.children.length !== 1 ||
    parent.children[0]!.id !== controlId
  )
    throw Error('Inline presence requires one exact native control revision.');
  const revision = nativePropertyAddress(parent);
  if (!revision) throw Error('Inline presence has no native revision address.');
  const key = propertyAddressKey(partName, revision);
  const sites = collectRevisionSites(part).filter((site) => {
    const address = nativePropertyAddress(site.node);
    return address && propertyAddressKey(partName, address) === key;
  });
  if (sites.length !== 1 || sites[0]!.node.id !== parent.id || sites[0]!.refused)
    throw Error('Inline presence has ambiguous or unsupported native membership.');
  if (
    expected &&
    (expected.blockId !== blockId ||
      expected.partName !== partName ||
      expected.kind !== kind ||
      propertyAddressKey(expected.partName, expected.revision) !== key)
  )
    throw Error('Inline presence native ownership changed.');
  return { blockId, partName, kind, revision, siteNodeId: parent.id };
}

export function inlinePresenceExtension(owners: readonly InlinePresenceNative[]): JsonObject {
  return {
    nativeInline: {
      schemaVersion: owners.some((owner) => owner.cell)
        ? 3
        : owners.some((owner) => owner.gap)
          ? 2
          : 1,
      owners: owners.map(({ blockId, partName, kind, revision, gap, cell }) => ({
        blockId,
        partName,
        kind,
        revision: { ...revision },
        ...(gap ? { gap: gap.map((address) => ({ ...address })) } : {}),
        ...(cell ? { cell: { kind: cell.kind, revision: { ...cell.revision } } } : {}),
      })),
    },
  };
}

/** Older application readers refuse these inline roots rather than approximating their decisions. */
export function readInlinePresenceExtension(
  record: BlockPresenceRecord
): ReadonlyMap<string, InlinePresenceNative> | null {
  const data = record.extension?.nativeInline;
  if (data === undefined) return null;
  if (record.schemaVersion !== 3 && record.schemaVersion !== 4 && record.schemaVersion !== 5)
    throw Error('Inline presence requires a supported private proposal schema.');
  if (
    !object(data) ||
    !exact(data, ['schemaVersion', 'owners']) ||
    data.schemaVersion !== (record.schemaVersion === 5 ? 3 : record.schemaVersion === 4 ? 2 : 1) ||
    !Array.isArray(data.owners) ||
    !data.owners.length ||
    data.owners.length > 1024
  )
    throw Error('Unsupported inline presence ownership profile.');
  const owners = new Map<string, InlinePresenceNative>();
  for (const entry of data.owners) {
    if (
      !object(entry) ||
      !exact(entry, [
        'blockId',
        'partName',
        'kind',
        'revision',
        ...(object(entry) && Object.hasOwn(entry, 'gap') ? ['gap'] : []),
        ...(object(entry) && Object.hasOwn(entry, 'cell') ? ['cell'] : []),
      ]) ||
      typeof entry.blockId !== 'string' ||
      !entry.blockId ||
      typeof entry.partName !== 'string' ||
      (entry.kind !== 'ins' && entry.kind !== 'del') ||
      !object(entry.revision) ||
      typeof entry.revision.id !== 'string' ||
      !entry.revision.id ||
      typeof entry.revision.author !== 'string' ||
      !entry.revision.author ||
      (entry.revision.date !== undefined && typeof entry.revision.date !== 'string') ||
      !Object.keys(entry.revision).every((key) => ['id', 'author', 'date'].includes(key)) ||
      owners.has(entry.blockId)
    )
      throw Error('Invalid inline presence native ownership.');
    if (Object.hasOwn(entry, 'gap')) {
      if (
        (record.schemaVersion !== 4 && record.schemaVersion !== 5) ||
        entry.kind !== 'ins' ||
        !Array.isArray(entry.gap) ||
        !entry.gap.length ||
        entry.gap.length > 1024
      )
        throw Error('Invalid paragraph birth native membership.');
      for (const address of entry.gap)
        if (
          !object(address) ||
          typeof address.id !== 'string' ||
          !address.id ||
          typeof address.author !== 'string' ||
          (!address.author && !(record.schemaVersion === 5 && entry.cell)) ||
          (address.date !== undefined && typeof address.date !== 'string') ||
          !Object.keys(address).every((key) => ['id', 'author', 'date'].includes(key))
        )
          throw Error('Invalid paragraph birth native address.');
    }
    if (
      Object.hasOwn(entry, 'cell') &&
      (record.schemaVersion !== 5 ||
        entry.kind !== 'ins' ||
        !entry.gap ||
        !object(entry.cell) ||
        !exact(entry.cell, ['kind', 'revision']) ||
        !['row', 'cell'].includes(String(entry.cell.kind)) ||
        !object(entry.cell.revision) ||
        !Object.keys(entry.cell.revision).every((key) => ['id', 'author', 'date'].includes(key)) ||
        typeof entry.cell.revision.id !== 'string' ||
        !entry.cell.revision.id ||
        typeof entry.cell.revision.author !== 'string' ||
        !entry.cell.revision.author ||
        (entry.cell.revision.date !== undefined && typeof entry.cell.revision.date !== 'string'))
    )
      throw Error('Invalid native cell ownership.');
    owners.set(entry.blockId, entry as unknown as InlinePresenceNative);
  }
  if (record.schemaVersion === 4 && ![...owners.values()].some((owner) => owner.gap))
    throw Error('Paragraph birth profile has no structural owner.');
  if (record.schemaVersion === 5 && ![...owners.values()].some((owner) => owner.cell))
    throw Error('Cell birth profile has no topology owner.');
  return owners;
}
