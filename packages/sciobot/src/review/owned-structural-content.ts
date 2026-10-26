import {
  canonicalOoxmlFingerprint,
  WML_NAMESPACE_URI,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';

/**
 * A marked paragraph's selected text revisions own their contents. Committed
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
    const children: OoxmlNode[] = [];
    for (const child of node.children) {
      if (independent && plain(child)) continue;
      const current = visit(child, depth + 1);
      const previous = children.at(-1);
      // A committed insertion can split an owned insertion around itself. Join only
      // its direct fragments with exactly the same native revision attributes.
      if (
        independent &&
        previous &&
        mark(previous) &&
        mark(current) &&
        previous.kind !== 'textValue' &&
        current.kind !== 'textValue' &&
        canonicalOoxmlFingerprint({ ...previous, children: [] } as OoxmlNode) ===
          canonicalOoxmlFingerprint({ ...current, children: [] } as OoxmlNode)
      ) {
        children[children.length - 1] = {
          ...previous,
          children: [...previous.children, ...current.children],
        } as OoxmlNode;
      } else children.push(current);
    }
    if (
      children.length === node.children.length &&
      children.every((child, i) => child === node.children[i])
    )
      return node;
    return { ...node, children } as OoxmlNode;
  };
  return visit(root, 0);
}
