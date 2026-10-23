import { preservesInsertedParagraphProperty } from './inserted-paragraph-property-presence.ts';
// A decided table structure must not discard an offline peer's future paragraph text.
import { transactionActorId } from '../package/actor-scoped-ids.ts';
import {
  createNodeIdAllocator,
  findNode,
  insertChildren,
  parentNodeOf,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { retireParagraph } from '../package/retired-paragraph.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlParagraphNode,
  type OoxmlParagraphPropertiesNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';

const CONTAINERS = new Set([
  'body',
  'tc',
  'sdtContent',
  'hdr',
  'ftr',
  'footnote',
  'endnote',
  'txbxContent',
  'customXml',
]);

/** Retain removed cell addresses at the nearest surviving story boundary, without their decided content. */
export function retainRemovedTableParagraphs(
  before: OoxmlPart,
  decided: ReadonlySet<string>,
  after: OoxmlPart,
  options?: EditOptions,
  resolveParagraph?: (paragraph: OoxmlParagraphNode) => OoxmlParagraphNode | undefined
): OoxmlPart | null {
  const shared = transactionActorId() !== undefined;
  const roots = new Map<string, OoxmlElement>();
  for (const id of decided) {
    const mark = findNode(before, id);
    const properties = parentNodeOf(before, id);
    if (
      !mark ||
      mark.kind === 'textValue' ||
      mark.namespaceUri !== WML_NAMESPACE_URI ||
      !properties
    )
      continue;
    const row = ['ins', 'del'].includes(mark.localName) && properties.localName === 'trPr';
    const cell = ['cellIns', 'cellDel'].includes(mark.localName) && properties.localName === 'tcPr';
    if (!row && !cell) continue;
    const owner = parentNodeOf(before, properties.id);
    if (owner && !findNode(after, owner.id)) roots.set(owner.id, owner);
  }
  const seen = new Set<string>();
  const mint = createNodeIdAllocator(before);
  let result = after;
  for (const root of roots.values()) {
    let anchor: OoxmlNode = root;
    let container = parentNodeOf(before, anchor.id);
    while (
      container &&
      !(
        container.namespaceUri === WML_NAMESPACE_URI &&
        CONTAINERS.has(container.localName) &&
        findNode(result, container.id)
      )
    ) {
      anchor = container;
      container = parentNodeOf(before, anchor.id);
    }
    if (!container) continue;
    const paragraphs: OoxmlParagraphNode[] = [];
    const visit = (node: OoxmlNode): void => {
      if (node.kind === 'textValue') return;
      if (node.kind === 'paragraph') {
        if (!seen.has(node.id) && !findNode(result, node.id)) {
          seen.add(node.id);
          const hasOwnedText = (child: OoxmlNode): boolean =>
            child.kind !== 'textValue' &&
            ((child.namespaceUri === WML_NAMESPACE_URI &&
              ['ins', 'del'].includes(child.localName) &&
              decided.has(child.id)) ||
              child.children.some(hasOwnedText));
          // An ordinary row owns all its content. Only independently marked paragraphs
          // retain unowned text when their enclosing structure disappears.
          if (!hasOwnedText(node)) return;
          const resolved = resolveParagraph?.(node);
          const survivor = resolved ?? {
            ...node,
            children: node.children.filter((child) => child.kind === 'paragraphProperties'),
          };
          const content = survivor.children.some((child) => child.kind !== 'paragraphProperties');
          if (!shared && !content) return;
          const retired = retireParagraph(withoutSpentRunHolder(node, survivor, decided), mint);
          paragraphs.push(retired);
        }
        return;
      }
      for (const child of node.children) visit(child);
    };
    visit(root);
    if (!paragraphs.length) continue;
    const current = findNode(result, container.id);
    if (!current || current.kind === 'textValue') return null;
    const at = container.children.findIndex((node) => node.id === anchor.id);
    const next = container.children
      .slice(at)
      .find((node) => current.children.some((child) => child.id === node.id));
    const index = next
      ? current.children.findIndex((node) => node.id === next.id)
      : current.children.findIndex(
          (node) =>
            node.kind !== 'textValue' &&
            node.namespaceUri === WML_NAMESPACE_URI &&
            node.localName === 'sectPr'
        );
    const inserted = insertChildren(
      result,
      current.id,
      index < 0 ? current.children.length : index,
      paragraphs,
      options
    );
    if (!inserted.ok) return null;
    result = inserted.part;
  }
  return result;
}

/** A removed insertion cannot leave its generated paragraph-mark holder on recovered text. */
function withoutSpentRunHolder(
  before: OoxmlParagraphNode,
  after: OoxmlParagraphNode,
  decided: ReadonlySet<string>
): OoxmlParagraphNode {
  return {
    ...after,
    children: after.children.map((properties) => {
      if (properties.kind !== 'paragraphProperties') return properties;
      const original = before.children.find((node) => node.id === properties.id);
      if (!original) return properties;
      return {
        ...properties,
        children: properties.children.filter((holder) => {
          if (
            holder.namespaceUri !== WML_NAMESPACE_URI ||
            holder.localName !== 'rPr' ||
            holder.children.length ||
            holder.attributes.length
          )
            return true;
          const prior = original.children.find((node) => node.id === holder.id);
          return (
            !prior ||
            prior.kind === 'textValue' ||
            !prior.children.length ||
            !prior.children.every(
              (mark) =>
                mark.kind !== 'textValue' &&
                mark.namespaceUri === WML_NAMESPACE_URI &&
                mark.localName === 'ins' &&
                decided.has(mark.id) &&
                !preservesInsertedParagraphProperty(mark, 'rPr')
            )
          );
        }),
      } as OoxmlParagraphPropertiesNode;
    }),
  };
}
