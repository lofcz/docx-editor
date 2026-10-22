import type { OoxmlNode, OoxmlParagraphNode } from '../package/ooxml-tree.ts';

export function contains(node: OoxmlNode, id: string): boolean {
  if (node.id === id) return true;
  if (node.kind === 'textValue') return false;
  return node.children.some((child) => contains(child, id));
}

/** Whether the element holding a text value is a `w:t` or a `w:delText`. */
export function textKindOfValue(root: OoxmlNode, valueId: string): 'text' | 'deletedText' {
  const walk = (node: OoxmlNode): 'text' | 'deletedText' | null => {
    if (node.kind === 'textValue') return null;
    if (
      (node.kind === 'text' || node.kind === 'deletedText') &&
      node.children.some((child) => child.id === valueId)
    ) {
      return node.kind;
    }
    for (const child of node.children) {
      const found = walk(child);
      if (found) return found;
    }
    return null;
  };
  return walk(root) ?? 'text';
}

export function findTextParent(paragraph: OoxmlParagraphNode, valueId: string): OoxmlNode | null {
  const walk = (node: OoxmlNode): OoxmlNode | null => {
    if (node.kind === 'textValue') return null;
    if (
      (node.kind === 'text' || node.kind === 'deletedText') &&
      node.children.some((child) => child.id === valueId)
    ) {
      return node;
    }
    for (const child of node.children) {
      const found = walk(child);
      if (found) return found;
    }
    return null;
  };
  return walk(paragraph);
}
