/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type * as Y from 'yjs';
import { TEXT_DELETIONS_KEY } from './paragraph-text-deletions.ts';
import { TEXT_DECISIONS_KEY } from './paragraph-text-decisions.ts';
import { asTrackedType } from './registry-node-reads.ts';
import type { PackageSchema } from './schema.ts';

/** Every shared container whose local changes participate in document history. */
export function registryTrackedTypes(
  schema: PackageSchema,
  doc: Y.Doc
): readonly Y.AbstractType<unknown>[] {
  return [
    asTrackedType(schema.meta),
    asTrackedType(schema.nodes),
    asTrackedType(schema.parts),
    asTrackedType(schema.relationships),
    asTrackedType(schema.overrides),
    asTrackedType(schema.defaults),
    asTrackedType(schema.binaries),
    asTrackedType(schema.attributes),
    asTrackedType(schema.bindings),
    // Withdrawing a deletion or review decision restores the affected text.
    asTrackedType(doc.getMap(TEXT_DELETIONS_KEY)),
    asTrackedType(doc.getMap(TEXT_DECISIONS_KEY)),
  ];
}
