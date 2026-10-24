import * as Y from 'yjs';
import {
  PROPERTY_GROUP_RECEIPTS,
  MAX_PROPERTY_GROUPS,
  decodePropertyGroupBinding,
  encodePropertyGroupBinding,
} from './property-group-codec';
import {
  MAX_REVIEW_FIELD_BYTES,
  MAX_REVIEW_FIELD_TARGETS,
  REVIEW_FIELD_INTENTS,
} from './review-field-codec';
import { MAX_REVIEW_FLAGS, REVIEW_FLAG_INTENTS } from './review-flag-codec';
import { assertNativePropertyProfile, propertyScopedReceipt } from './property-review-scope';
import type { PropertyReviewReceipt } from './property-review-prototype';
import type { PropertyGroupScope } from '../review/property-history-contract';

/** Application-owned binding records share their proposal's original native history frame. */
export function createPropertyGroupBindings(doc: Y.Doc) {
  const map = doc.getMap<string>(PROPERTY_GROUP_RECEIPTS);
  const origins = new Set<unknown>();
  const records = () => {
    if (map._map.size > MAX_PROPERTY_GROUPS)
      throw new Error('Property group bindings exceed their record limit.');
    return [...map].map(([id, raw]) => decodePropertyGroupBinding(id, raw));
  };
  const before = (transaction: Y.Transaction) => {
    const manager = transaction.origin;
    if (
      manager instanceof Y.UndoManager &&
      [...origins].some((origin) => manager.trackedOrigins.has(origin))
    )
      manager.addToScope(map);
  };
  doc.on('beforeTransaction', before);
  return Object.freeze({
    assertHealthy() {
      if (!doc.store.pendingStructs && !doc.store.pendingDs) records();
    },
    reserve(scopes: readonly PropertyGroupScope[], fieldWrites: number, flagWrites: number) {
      records();
      if (
        !Number.isSafeInteger(fieldWrites) ||
        fieldWrites < 0 ||
        !Number.isSafeInteger(flagWrites) ||
        flagWrites < 0
      )
        throw new Error('Property group bindings require complete bounded write budgets.');
      if (doc.store.pendingStructs || doc.store.pendingDs)
        throw new Error('Property group bindings need complete room updates.');
      if (map._map.size + scopes.length > MAX_PROPERTY_GROUPS)
        throw new Error('Property group bindings exceed their record limit.');
      if (scopes.length) assertNativePropertyProfile(doc);
      for (const scope of scopes) {
        if (map.has(scope.groupId))
          throw new Error('Property group binding identity is already reserved.');
        const memberBytes = new TextEncoder().encode(JSON.stringify(scope.membersKey)).length;
        // A client:clock address has at most 33 decimal characters and three JSON bytes.
        const forecast =
          512 +
          memberBytes +
          36 *
            (Math.min(fieldWrites, MAX_REVIEW_FIELD_TARGETS) +
              Math.min(flagWrites, MAX_REVIEW_FLAGS));
        if (forecast > MAX_REVIEW_FIELD_BYTES)
          throw new Error('Property group binding exceeds its complete byte budget.');
      }
    },
    publish(scopes: readonly PropertyGroupScope[], receipt: PropertyReviewReceipt) {
      if (!doc._transaction?.local)
        throw new Error('Property group bindings require native publication.');
      const old = records();
      const claimedFields = new Set(old.flatMap((record) => record.fields));
      const claimedFlags = new Set(old.flatMap((record) => record.flags));
      const staged: [string, string][] = [];
      for (const scope of scopes) {
        const selected = propertyScopedReceipt(doc, receipt, scope.nodeIds);
        for (const item of selected.fields) {
          if (claimedFields.has(item.id))
            throw new Error('Property field ownership belongs to another group.');
          claimedFields.add(item.id);
        }
        for (const item of selected.flags) {
          if (claimedFlags.has(item.id))
            throw new Error('Property flag ownership belongs to another group.');
          claimedFlags.add(item.id);
        }
        const raw = encodePropertyGroupBinding({
          schemaVersion: 1,
          groupId: scope.groupId,
          membersKey: scope.membersKey,
          fields: selected.fields.map((item) => item.id),
          flags: selected.flags.map((item) => item.id),
        });
        if (map.has(scope.groupId)) throw new Error('Property group binding is immutable.');
        staged.push([scope.groupId, raw]);
      }
      if (map._map.size + staged.length > MAX_PROPERTY_GROUPS)
        throw new Error('Property group bindings exceed their record limit.');
      origins.add(doc._transaction.origin);
      for (const [id, raw] of staged) map.set(id, raw);
    },
    read(groupId: string, membersKey: string): PropertyReviewReceipt | null {
      records();
      const raw = map.get(groupId);
      if (raw === undefined) return null;
      assertNativePropertyProfile(doc);
      const binding = decodePropertyGroupBinding(groupId, raw);
      if (binding.membersKey !== membersKey)
        throw new Error('Property group membership changed after ownership capture.');
      const resolve = (ids: readonly string[], name: string) =>
        ids.map((id) => {
          const record = doc.getMap(name).get(id);
          if (typeof record !== 'string') throw new Error('Property group intent is missing.');
          return Object.freeze({ id, record });
        });
      return Object.freeze({
        fields: Object.freeze(resolve(binding.fields, REVIEW_FIELD_INTENTS)),
        flags: Object.freeze(resolve(binding.flags, REVIEW_FLAG_INTENTS)),
      });
    },
    dispose() {
      doc.off('beforeTransaction', before);
    },
  });
}
