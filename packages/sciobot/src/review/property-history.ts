import type { EditorCollaborationSession } from '@docx-editor.dev/core/collaboration';
import type {
  PropertyReviewPrototype,
  PropertyReviewReceipt,
} from '../collaboration/property-review-prototype';
import { PROPERTY_REVIEW_PROTOTYPE, propertyGroupMembersKey } from './property-history-contract';
import type { NativeReviewGroup } from './groups';

/** Internal capability only. Default production attachment returns no property ownership scope. */
export function propertyHistoryOf(
  session: EditorCollaborationSession | null
): PropertyReviewPrototype | null {
  if (!session) return null;
  const scope = (session as unknown as Record<symbol, unknown>)[PROPERTY_REVIEW_PROTOTYPE];
  if (scope === undefined || scope === null) return null;
  if (
    typeof scope !== 'object' ||
    typeof (scope as PropertyReviewPrototype).read !== 'function' ||
    typeof (scope as PropertyReviewPrototype).withdrawDuringPublication !== 'function'
  )
    throw new Error('Invalid native property history capability.');
  return scope as PropertyReviewPrototype;
}
export function readPropertyGroupReceipt(
  scope: PropertyReviewPrototype,
  group: NativeReviewGroup
): PropertyReviewReceipt | null {
  if (group.properties.length || !group.native.length || group.readOnly) return null;
  return scope.read(group.id, propertyGroupMembersKey(group));
}
export function combinePropertyGroupReceipts(
  receipts: readonly PropertyReviewReceipt[]
): PropertyReviewReceipt {
  const fields = new Map<string, PropertyReviewReceipt['fields'][number]>();
  const flags = new Map<string, PropertyReviewReceipt['flags'][number]>();
  for (const receipt of receipts) {
    for (const item of receipt.fields) {
      if (fields.has(item.id)) throw new Error('Selected property groups share field ownership.');
      fields.set(item.id, item);
    }
    for (const item of receipt.flags) {
      if (flags.has(item.id)) throw new Error('Selected property groups share flag ownership.');
      flags.set(item.id, item);
    }
  }
  return Object.freeze({
    fields: Object.freeze([...fields.values()]),
    flags: Object.freeze([...flags.values()]),
  });
}
