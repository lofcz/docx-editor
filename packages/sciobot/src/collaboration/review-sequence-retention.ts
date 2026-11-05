import * as Y from 'yjs';
import { hasEvaluatedNativeProfile } from './native-profile';

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
  if (typeof value !== 'string' || !hasEvaluatedNativeProfile(doc)) return false;
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
