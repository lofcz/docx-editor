import {
  parentNodeOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';

const PROPERTY_RECORDS = new Set([
  'rPrChange',
  'pPrChange',
  'tblPrChange',
  'trPrChange',
  'tcPrChange',
  'tblGridChange',
]);

/** Current property bag that rejecting this record can overwrite. A read, with no native preview. */
export function currentRevisionPropertyContent(part: OoxmlPart, node: OoxmlNode): OoxmlNode | null {
  if (
    node.kind === 'textValue' ||
    node.namespaceUri !== WML_NAMESPACE_URI ||
    !PROPERTY_RECORDS.has(node.localName)
  )
    return null;
  const parent = parentNodeOf(part, node.id);
  if (!parent) throw new Error('An owned property record has no native property container.');
  return {
    ...parent,
    children: parent.children.filter((child) => child.id !== node.id),
  } as OoxmlNode;
}

/** Current content that an inverse can replace, including a whole paragraph owned by a mark. */
export function currentRevisionGuardContent(part: OoxmlPart, node: OoxmlNode): OoxmlNode | null {
  const properties = currentRevisionPropertyContent(part, node);
  if (properties) return properties;
  if (node.kind === 'textValue' || node.namespaceUri !== WML_NAMESPACE_URI) return null;
  const holder = parentNodeOf(part, node.id);
  if (!holder || holder.namespaceUri !== WML_NAMESPACE_URI) return null;
  if (['cellIns', 'cellDel'].includes(node.localName) && holder.localName === 'tcPr') {
    const cell = parentNodeOf(part, holder.id);
    return cell?.kind === 'tableCell' ? cell : null;
  }
  if (!['ins', 'del'].includes(node.localName)) return null;
  if (holder.localName === 'trPr') {
    const row = parentNodeOf(part, holder.id);
    return row?.kind === 'tableRow' ? row : null;
  }
  if (holder.localName !== 'rPr') return null;
  const paragraphProperties = parentNodeOf(part, holder.id);
  if (
    paragraphProperties?.namespaceUri !== WML_NAMESPACE_URI ||
    paragraphProperties.localName !== 'pPr'
  )
    return null;
  const paragraph = parentNodeOf(part, paragraphProperties.id);
  return paragraph?.kind === 'paragraph' ? paragraph : null;
}
