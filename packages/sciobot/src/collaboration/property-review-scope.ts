import { independentPropertyChild } from '../review/property-boundaries';
import * as Y from 'yjs';
import { decodeFieldIntent, decodeFieldTarget, REVIEW_FIELD_TARGETS } from './review-field-codec';
import { decodeFlagTarget } from './review-flag-codec';
import { reviewArray } from './review-sequences';
import { flagMap } from './review-flag-items';
import type { PropertyReviewReceipt } from './property-review-prototype';

/** Fail before authoring against an unknown native property representation. */
export function assertNativePropertyProfile(doc: Y.Doc): void {
  const meta = doc.share.has('docx-package-meta-v1') ? doc.getMap('docx-package-meta-v1') : null;
  if (
    !meta ||
    meta.get('protocolVersion') !== 1 ||
    meta.get('sharedSchemaVersion') !== 9 ||
    meta.get('repairVersion') !== 14 ||
    meta.get('canonicalModelVersion') !== 1
  )
    throw new Error('Unsupported native property ownership profile.');
}

/** Pinned native wire selection only. Canonical property IDs come from both source and proposal. */
export function propertyScopedReceipt(
  doc: Y.Doc,
  receipt: PropertyReviewReceipt,
  nodeIds: ReadonlySet<string>
): PropertyReviewReceipt {
  assertNativePropertyProfile(doc);
  const nativeNodes = doc.share.has('docx-package-nodes-v1')
    ? doc.getMap('docx-package-nodes-v1')
    : null;
  if (!nativeNodes || !nodeIds.size)
    throw new Error('Property ownership requires native node identities.');
  const wordName = (record: unknown): string | null => {
    if (!(record instanceof Y.Map)) return null;
    const packed = record.get('s') ?? record.get('b');
    if (typeof packed !== 'string') return null;
    const parts = packed.split('\u001f');
    return parts.length === 4 && parts[1] === 'w' ? parts[2]! : null;
  };
  const independent = (parent: unknown, childId: string | null) => {
    const parentName = wordName(parent),
      childName = childId === null ? null : wordName(nativeNodes.get(childId));
    return (
      parentName !== null && childName !== null && independentPropertyChild(parentName, childName)
    );
  };
  // The native publisher replaces temporary #new IDs with logical wire IDs.
  // Resolve the complete property subtree from its stable existing containers.
  const resolvedIds = new Set(nodeIds),
    pending = [...nodeIds];
  while (pending.length) {
    const record = nativeNodes.get(pending.pop()!);
    if (!(record instanceof Y.Map)) continue;
    const children = record.get('children');
    if (!(children instanceof Y.Array)) continue;
    for (const child of children) {
      if (typeof child !== 'string') throw new Error('Invalid native property child identity.');
      if (independent(record, child)) continue;
      if (!resolvedIds.has(child)) {
        resolvedIds.add(child);
        pending.push(child);
      }
    }
  }
  const ownsMap = (map: Y.Map<unknown> | null) =>
    map?._item?.parent === nativeNodes &&
    typeof map._item.parentSub === 'string' &&
    resolvedIds.has(map._item.parentSub);
  const fields = receipt.fields.filter((entry) => {
    const intent = decodeFieldIntent(entry.id, entry.record);
    const target = decodeFieldTarget(
      intent.target,
      doc.getMap(REVIEW_FIELD_TARGETS).get(intent.target)
    );
    if (target.sequence) {
      const array = reviewArray(doc, target),
        item = array?._item;
      return (
        item?.parentSub === 'children' &&
        item.parent instanceof Y.Map &&
        item.parent._item?.parent === nativeNodes &&
        !independent(item.parent, target.baseline) &&
        (ownsMap(item.parent) || resolvedIds.has(target.baseline!))
      );
    }
    if (
      target.root === 'docx-package-attributes-v1' ||
      target.root === 'docx-package-bindings-v1'
    ) {
      return resolvedIds.has(target.key.split('\u001f')[0]!);
    }
    if (target.root !== undefined) return false;
    return ownsMap(
      flagMap(doc, { ...target, client: target.client!, clock: target.clock!, baseline: 'missing' })
    );
  });
  const flags = receipt.flags.filter((entry) => {
    const target = decodeFlagTarget(
      entry.record,
      doc.getMap('sciobot-review-flag-targets-v1').get(entry.record)
    );
    return ownsMap(flagMap(doc, target));
  });
  return Object.freeze({ fields: Object.freeze(fields), flags: Object.freeze(flags) });
}
