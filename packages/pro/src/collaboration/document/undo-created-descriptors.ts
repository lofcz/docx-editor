/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type * as Y from 'yjs';
import type { DocumentRegistry } from './registry.ts';
import { creationsBeingUndone } from './undo-delete-filter.ts';
import { isElementRecord, parseAttributeMapKey, parseBindingMapKey } from './schema.ts';
import type { LogicalId } from './identity.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * Whether an undo keeps this attribute or binding of a node that its step created.
 *
 * A peer that lists the node again at the same time, as a concurrent decision or move does,
 * shows it whole only if the step's descriptors stay. This fork keeps them for WordprocessingML
 * nodes only. A birth that the undo marks undone keeps its own identity and geometry rules,
 * and a custom XML record must leave its part whole or not at all, so neither keeps more here:
 * retained attributes there read as content that keeps the node shown.
 */
export function retainsCreatedDescriptor(
  registry: DocumentRegistry,
  births: { covers(id: LogicalId): boolean },
  item: Y.Item
): boolean {
  if (item.parentSub === null) return false;
  const schema = registry.schema;
  const key =
    item.parent === schema.attributes
      ? parseAttributeMapKey(item.parentSub)
      : item.parent === schema.bindings
        ? parseBindingMapKey(item.parentSub)
        : null;
  if (!key || !creationsBeingUndone(schema)?.has(key.logicalId)) return false;
  const record = registry.record(key.logicalId);
  return (
    !!record &&
    isElementRecord(record) &&
    record.namespaceUri === W &&
    !births.covers(key.logicalId)
  );
}
