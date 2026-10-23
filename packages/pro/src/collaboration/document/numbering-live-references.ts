/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { DocumentRegistry } from './registry.ts';
import type { LogicalId } from './identity.ts';
import { isElementRecord, type ElementRecord } from './schema.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const value = (record: ElementRecord, name: string) =>
  record.attributes.find(
    (attribute) => attribute.namespaceUri === W && attribute.localName === name
  )?.value;

/** Definitions needed by live paragraph/style references, including previously undone resources. */
export function neededNumberingMembers(registry: DocumentRegistry): ReadonlySet<LogicalId> {
  const candidates = registry.allLogicalIds().flatMap((id) => {
    const record = registry.record(id);
    return record &&
      isElementRecord(record) &&
      record.namespaceUri === W &&
      ['numId', 'num', 'abstractNum'].includes(record.localName) &&
      !registry.isTombstoned(id)
      ? [record]
      : [];
  });
  const roots = new Set(registry.partEntries().map((part) => part.rootLogicalId));
  const live = (id: LogicalId) => {
    const seen = new Set<LogicalId>();
    let current: LogicalId | null = id;
    while (current !== null && !seen.has(current)) {
      if (registry.isTombstoned(current)) return false;
      if (roots.has(current)) return true;
      seen.add(current);
      current = registry.parentOf(current);
    }
    return false;
  };
  const referenced = new Set(
    candidates
      .filter((node) => node.localName === 'numId' && live(node.logicalId))
      .map((node) => value(node, 'val'))
  );
  const needed = new Set<LogicalId>();
  const abstractIds = new Set<string | undefined>();
  for (const num of candidates) {
    if (num.localName !== 'num' || !referenced.has(value(num, 'numId'))) continue;
    needed.add(num.logicalId);
    const reference = num.childIds
      .map((id) => registry.record(id))
      .find((node) => node && isElementRecord(node) && node.localName === 'abstractNumId');
    if (reference && isElementRecord(reference)) abstractIds.add(value(reference, 'val'));
  }
  for (const abstract of candidates)
    if (abstract.localName === 'abstractNum' && abstractIds.has(value(abstract, 'abstractNumId')))
      needed.add(abstract.logicalId);
  return needed;
}
