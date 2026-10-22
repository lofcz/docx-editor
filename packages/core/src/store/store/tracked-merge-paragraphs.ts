import { isHiddenTableParagraph } from './table-transaction-state.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlParagraphNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { parentNodeOf } from '../package/ooxml-edit.ts';
import type { TransactionRevisionIds } from '../package/ooxml-edit.ts';
import { MAX_XML_DEPTH } from '../package/xml-reader.ts';
import { revisionKey } from './tree-op-tracked-adjacency.ts';
import { paragraphIsBlank } from './tree-op-table-cell-shape.ts';

/** Resolve only inline wrappers this transaction wrote, before making a new tracked merge copy. */
export function trackedMergeParagraphs(
  part: OoxmlPart,
  nodes: readonly OoxmlNode[],
  owned: TransactionRevisionIds | undefined
): readonly OoxmlParagraphNode[] | null {
  const project = (node: OoxmlNode, depth: number): readonly OoxmlNode[] | null => {
    if (depth > MAX_XML_DEPTH) return null;
    if (node.kind === 'textValue') return [node];
    const word = node.namespaceUri === WML_NAMESPACE_URI;
    if (
      word &&
      (/PrChange$/.test(node.localName) ||
        node.localName === 'tblGridChange' ||
        ['moveFrom', 'moveTo'].includes(node.localName))
    )
      return null;
    const wrapper = word && ['ins', 'del'].includes(node.localName);
    if (wrapper) {
      const attr = (name: string) =>
        node.attributes.find(
          (value) => value.namespaceUri === WML_NAMESPACE_URI && value.localName === name
        )?.value;
      const id = attr('id'),
        author = attr('author');
      if (
        id === undefined ||
        author === undefined ||
        !owned?.wroteUnder(revisionKey(id, author, attr('date')))
      )
        return null;
    }
    const children: OoxmlNode[] = [];
    for (const child of node.children) {
      const projected = project(child, depth + 1);
      if (!projected) return null;
      for (const value of projected) children.push(value);
    }
    if (wrapper) return node.localName === 'del' ? [] : children;
    return [{ ...node, children } as OoxmlNode];
  };
  const result: OoxmlParagraphNode[] = [];
  for (const node of nodes) {
    if (isHiddenTableParagraph(owned, node.id)) continue;
    if (node.kind !== 'paragraph') return null;
    const projected = project(node, 0);
    if (!projected || projected.length !== 1 || projected[0]?.kind !== 'paragraph') return null;
    const owner = parentNodeOf(part, node.id);
    const onlyBlock =
      owner?.kind === 'tableCell' &&
      owner.children.filter(
        (child) => child.namespaceUri !== WML_NAMESPACE_URI || child.localName !== 'tcPr'
      ).length === 1;
    if (!onlyBlock || !paragraphIsBlank(projected[0])) result.push(projected[0]);
  }
  return result;
}
