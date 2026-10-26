import * as Y from 'yjs';

/**
 * The document transport retains undone birth ancestry for foreign content. Sequence
 * ownership withdraws the author's insertion, while that transport owns its visibility.
 * Read only this pinned wire profile; no transport implementation enters production.
 */
export function retainedBirthSequenceEntry(
  doc: Y.Doc,
  array: Y.Array<unknown>,
  value: string | null
): boolean {
  const existingMap = (name: string): Y.Map<unknown> | null => {
    if (!doc.share.has(name)) return null;
    // A cold update can arrive before its provider declares root constructors. Declaring
    // an existing root changes no shared data and does not create an absent directory.
    try {
      return doc.getMap(name);
    } catch {
      return null;
    }
  };
  const meta = existingMap('docx-package-meta-v1');
  if (!meta || typeof value !== 'string') return false;
  if (
    meta.get('protocolVersion') !== 1 ||
    meta.get('sharedSchemaVersion') !== 9 ||
    meta.get('repairVersion') !== 15 ||
    meta.get('canonicalModelVersion') !== 1
  )
    return false;
  const nodes = existingMap('docx-package-nodes-v1');
  if (!nodes) return false;
  const field = array._item;
  const parent = field?.parent;
  if (
    !field ||
    field.deleted ||
    field.parentSub !== 'children' ||
    !(parent instanceof Y.Map) ||
    parent.get('children') !== array ||
    parent._item?.parent !== nodes
  )
    return false;
  const record = nodes.get(value);
  return record instanceof Y.Map && record.get('undone') === true && record.get('deleted') !== true;
}
