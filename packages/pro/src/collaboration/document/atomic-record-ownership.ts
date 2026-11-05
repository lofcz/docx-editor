/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { ATOMIC_CONTENT_NAMESPACE } from '@docx-editor.dev/core/store';
import type { DocumentRegistry } from './registry.ts';
import type { LogicalId } from './identity.ts';
import type { SharedEffect } from './shared-effect.ts';

/** Resolve ownership against the transaction's final parent graph, including new controls. */
export function atomicRecordOwnership(
  registry: DocumentRegistry,
  effects: readonly SharedEffect[]
): ReadonlySet<LogicalId> {
  const parents = new Map<LogicalId, LogicalId>();
  const candidates = new Set<LogicalId>();
  const markers = new Map<LogicalId, boolean>();
  const kinds = new Map<LogicalId, string>();
  for (const effect of effects) {
    if ('logicalId' in effect) candidates.add(effect.logicalId);
    if (effect.kind === 'putNode') {
      candidates.add(effect.descriptor.logicalId);
      kinds.set(effect.descriptor.logicalId, effect.descriptor.kind);
    }
    if (
      effect.kind === 'setAttribute' &&
      effect.qname.namespaceUri === ATOMIC_CONTENT_NAMESPACE &&
      effect.qname.localName === 'atomicContent'
    )
      markers.set(effect.logicalId, effect.value === '1');
    if (effect.kind === 'spliceChildren') {
      candidates.add(effect.parentLogicalId);
      for (const child of effect.childLogicalIds) {
        parents.set(child, effect.parentLogicalId);
        candidates.add(child);
      }
    }
    if (effect.kind === 'moveNode')
      parents.set(effect.logicalId, effect.destinationParentLogicalId);
  }
  const marked = (id: LogicalId): boolean => {
    if ((kinds.get(id) ?? registry.kindOf(id)) !== 'contentControl') return false;
    if (markers.has(id)) return markers.get(id)!;
    const record = registry.record(id);
    return (
      record?.kind === 'contentControl' &&
      'attributes' in record &&
      record.attributes.some(
        (attribute) =>
          attribute.namespaceUri === ATOMIC_CONTENT_NAMESPACE &&
          attribute.localName === 'atomicContent' &&
          attribute.value === '1'
      )
    );
  };
  const owned = new Set<LogicalId>();
  for (const candidate of candidates) {
    let current: LogicalId | null = candidate;
    const seen = new Set<LogicalId>();
    while (current && seen.size <= registry.limits.maxTreeDepth && !seen.has(current)) {
      seen.add(current);
      if (marked(current)) {
        for (const id of seen) owned.add(id);
        break;
      }
      current = parents.get(current) ?? registry.parentOf(current);
    }
  }
  return owned;
}
