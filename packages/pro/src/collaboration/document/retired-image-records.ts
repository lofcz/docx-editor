/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import {
  IMAGE_RELATIONSHIP_TYPE,
  partNameKey,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { CanonicalBinaryDescriptor } from '@docx-editor.dev/core/collaboration/replication';
import type { DocumentRegistry } from './registry.ts';
import { DIGEST_PATTERN, rejectDangerousKey, rejectPartName, rejectString } from './limits.ts';
import { isNodeMap, isElementRecord, type EncodedRelationship } from './schema.ts';
import { asLogicalId } from './identity.ts';
import { freezeElement, freezeText } from './materialize-freeze.ts';
import { relsOwnerOf } from './materialize-rels.ts';
export const storageKeyOf = (name: string): string => partNameKey('/' + name.replace(/^\//, ''));
export function retiredImageRecords(registry: DocumentRegistry) {
  const binaries = new Map<
    string,
    CanonicalBinaryDescriptor & { readonly retainedOverride: boolean }
  >();
  registry.schema.binaries.forEach((value, key) => {
    if (!isNodeMap(value) || value.get('digest') !== '') return;
    const digest = value.get('retainedDigest'),
      size = value.get('size'),
      mediaType = value.get('mediaType'),
      storageKey = value.get('storageKey');
    if (
      typeof digest !== 'string' ||
      !DIGEST_PATTERN.test(digest) ||
      typeof size !== 'number' ||
      !Number.isSafeInteger(size) ||
      size < 0 ||
      size > registry.limits.maxBlobBytes ||
      typeof mediaType !== 'string' ||
      rejectString(mediaType, registry.limits.maxMediaTypeLength) ||
      typeof storageKey !== 'string' ||
      rejectString(storageKey, registry.limits.maxStorageKeyLength) ||
      rejectPartName('/' + storageKey.replace(/^\//, '')) ||
      storageKeyOf(key) !== storageKeyOf(storageKey)
    )
      return;
    binaries.set(storageKeyOf(storageKey), {
      digest,
      size,
      mediaType,
      storageKey,
      retainedOverride: value.get('retainedOverride') === true,
    });
  });
  const relationships: EncodedRelationship[] = [];
  registry.schema.relationships.forEach((holder) => {
    if (!isNodeMap(holder)) return;
    holder.forEach((value, id) => {
      if (
        !isNodeMap(value) ||
        value.get('id') !== '' ||
        value.get('type') !== IMAGE_RELATIONSHIP_TYPE ||
        value.get('targetMode') !== 'Internal' ||
        rejectString(id, registry.limits.maxStringLength)
      )
        return;
      const ownerPart = value.get('ownerPart'),
        rawTarget = value.get('rawTarget'),
        order = value.get('order');
      if (
        typeof ownerPart !== 'string' ||
        rejectPartName(ownerPart) ||
        typeof rawTarget !== 'string' ||
        rejectString(rawTarget, registry.limits.maxStringLength)
      )
        return;
      relationships.push({
        ownerPart,
        id,
        type: IMAGE_RELATIONSHIP_TYPE,
        rawTarget,
        targetMode: 'Internal',
        order: typeof order === 'number' && Number.isSafeInteger(order) ? order : 0,
      });
    });
  });
  return { binaries, relationships };
}
/** Read retained relationship XML through its original part root, including unknown extensions. */
export function retainedRelationshipsPart(
  registry: DocumentRegistry,
  name: string
): OoxmlPart | undefined {
  if (relsOwnerOf(name) === null || rejectPartName(name)) return undefined;
  const entry = registry.schema.parts.get(name);
  if (!isNodeMap(entry)) return undefined;
  const rootId = entry.get('retainedRootId');
  if (typeof rootId !== 'string' || !rootId || rejectDangerousKey(rootId)) return undefined;
  const seen = new Set<string>();
  let visited = 0;
  const read = (id: string, depth: number): OoxmlNode | null => {
    if (
      seen.has(id) ||
      depth > registry.limits.maxTreeDepth ||
      ++visited > registry.limits.maxNodes
    )
      return null;
    seen.add(id);
    const record = registry.record(asLogicalId(id));
    if (!record) return null;
    if (!isElementRecord(record)) return freezeText(record.logicalId, record.value);
    const children: OoxmlNode[] = [];
    for (const childId of record.childIds) {
      const child = read(childId, depth + 1);
      if (!child) return null;
      children.push(child);
    }
    return freezeElement(record, children);
  };
  const root = read(rootId, 0);
  if (
    !root ||
    root.kind === 'textValue' ||
    root.localName !== 'Relationships' ||
    root.namespaceUri !== 'http://schemas.openxmlformats.org/package/2006/relationships'
  )
    return undefined;
  return Object.freeze({
    id: name,
    name,
    contentType: 'application/vnd.openxmlformats-package.relationships+xml',
    root,
  });
}
