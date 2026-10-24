import { WML_NAMESPACE_URI, type OoxmlNode } from '@docx-editor.dev/core/store';

/**
 * A marked paragraph's selected text revisions own their contents. Trailing committed
 * text runs outside them survive the native inverse and are not part of that ownership.
 * Preserve paragraph properties, controls, unknown elements, and unselected revisions.
 */
export function ownedStructuralContent(root: OoxmlNode, selected: ReadonlySet<string>): OoxmlNode {
  const mark = (node: OoxmlNode): boolean =>
    node.kind !== 'textValue' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    node.localName === 'ins' &&
    selected.has(node.id);
  const revision = (node: OoxmlNode): boolean =>
    node.kind !== 'textValue' &&
    ((node.namespaceUri === WML_NAMESPACE_URI &&
      (['ins', 'del', 'moveFrom', 'moveTo'].includes(node.localName) ||
        node.localName.endsWith('Change'))) ||
      node.children.some(revision));
  const plain = (node: OoxmlNode): boolean =>
    node.kind === 'run' &&
    node.children.every(
      (child) => (child.kind === 'runProperties' && !revision(child)) || child.kind === 'text'
    );
  const visit = (node: OoxmlNode, depth: number): OoxmlNode => {
    if (depth > 256) throw Error('Owned structural content exceeds its depth limit.');
    if (node.kind === 'textValue') return node;
    const independent =
      node.kind === 'paragraph' &&
      node.children.some(
        (props) =>
          props.kind === 'paragraphProperties' &&
          props.children.some(
            (holder) => holder.kind === 'runProperties' && holder.children.some(mark)
          )
      ) &&
      node.children.some(mark);
    let end = node.children.length;
    if (independent) while (end > 0 && plain(node.children[end - 1]!)) end--;
    const children = node.children.slice(0, end).map((child) => visit(child, depth + 1));
    if (
      children.length === node.children.length &&
      children.every((child, i) => child === node.children[i])
    )
      return node;
    return { ...node, children } as OoxmlNode;
  };
  return visit(root, 0);
}
