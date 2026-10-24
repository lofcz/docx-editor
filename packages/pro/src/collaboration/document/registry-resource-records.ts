/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import * as Y from 'yjs';
import type { CanonicalBinaryDescriptor } from '@docx-editor.dev/core/collaboration/replication';
import {
  isNodeMap,
  makeBinaryEntry,
  makeRelationshipEntry,
  type EncodedRelationship,
  type PackageSchema,
} from './schema.ts';
import { relationshipKey } from './relationship-store.ts';
import { rejectDangerousKey } from './limits.ts';

/** Ordinary puts retain atomic descriptor replacement. Retirement keeps existing slots. */
export function putResourceRelationship(schema: PackageSchema, record: EncodedRelationship): void {
  if (rejectDangerousKey(record.ownerPart) || rejectDangerousKey(record.id)) return;
  const holder = new Y.Map<Y.Map<unknown>>();
  holder.set(record.id, makeRelationshipEntry(record));
  schema.relationships.set(relationshipKey(record.ownerPart, record.id), holder);
}
export function retireResourceRelationship(
  schema: PackageSchema,
  ownerPart: string,
  relationshipId: string
): void {
  for (const holder of [
    schema.relationships.get(relationshipKey(ownerPart, relationshipId)),
    schema.relationships.get(ownerPart),
  ]) {
    const entry = holder?.get(relationshipId);
    // Every existing reader omits an entry whose public relationship identity is empty.
    if (isNodeMap(entry)) entry.set('id', '');
  }
}
export function putResourceBinary(
  schema: PackageSchema,
  descriptor: CanonicalBinaryDescriptor,
  retainedOverride = true
): void {
  schema.binaries.set(descriptor.storageKey, makeBinaryEntry(descriptor, retainedOverride));
}
export function retireResourceBinary(schema: PackageSchema, storageKey: string): void {
  // Map-deleting one shared entry collapses concurrent removals. Undoing either then revives
  // the other's retired resource. Empty digests already mean "not materialized" to readers;
  // assigning that scalar retains both authors' independently reversible writes.
  const entry = schema.binaries.get(storageKey);
  if (isNodeMap(entry)) entry.set('digest', '');
}
