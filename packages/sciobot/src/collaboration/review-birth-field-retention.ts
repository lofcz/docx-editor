import * as Y from 'yjs';
import type { ReviewFieldLocation } from './review-field-codec';

const GEOMETRY = new Set([
  'gridCol',
  'gridSpan',
  'vMerge',
  'tcW',
  'tblW',
  'gridBefore',
  'gridAfter',
  'cantSplit',
  'trHeight',
  'tblHeader',
  'wBefore',
  'wAfter',
]);
/** Pinned wire interoperability only. The native provider owns these initial birth scalars. */
export function nativeBirthFieldCreation(
  doc: Y.Doc,
  transaction: Y.Transaction,
  target: ReviewFieldLocation,
  numberingConstructors = false
): boolean {
  if (target.sequence || target.root !== 'docx-package-attributes-v1') return false;
  const existing = (name: string): Y.Map<unknown> | null => {
    if (!doc.share.has(name)) return null;
    try {
      return doc.getMap(name);
    } catch {
      return null;
    }
  };
  const meta = existing('docx-package-meta-v1');
  if (
    !meta ||
    meta.get('protocolVersion') !== 1 ||
    meta.get('sharedSchemaVersion') !== 9 ||
    meta.get('repairVersion') !== 15 ||
    meta.get('canonicalModelVersion') !== 1
  )
    return false;
  const parts = target.key.split('\u001f');
  if (parts.length !== 3) return false;
  const [id, namespace, name] = parts;
  const nodes = existing('docx-package-nodes-v1'),
    record = nodes?.get(id!);
  if (
    !(record instanceof Y.Map) ||
    !record._item ||
    record._item.deleted ||
    record._item.id.clock < (transaction.beforeState.get(record._item.id.client) ?? 0)
  )
    return false;
  if (namespace === 'w14' && ['paraId', 'textId'].includes(name!)) return true;
  const shell = record.get('s') ?? record.get('b');
  if (typeof shell !== 'string') return false;
  const fields = shell.split('\u001f');
  return (
    fields.length === 4 &&
    fields[1] === 'w' &&
    (GEOMETRY.has(fields[2]!) || (numberingConstructors && ['numId', 'ilvl'].includes(fields[2]!)))
  );
}
