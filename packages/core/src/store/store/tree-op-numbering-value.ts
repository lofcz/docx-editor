import { WML_NAMESPACE_URI, type OoxmlNode } from '../package/ooxml-tree.ts';
import { build } from './tree-op-tracked-builders.ts';

/** Keep native numbering identity and opaque metadata when only a value changes. */
export function withNumberingValue(
  node: OoxmlNode,
  name: 'ilvl' | 'numId',
  value: string,
  nextId: () => string
): OoxmlNode {
  if (node.kind === 'textValue') return node;
  let found = false;
  const children = node.children.map((child) => {
    if (
      child.kind === 'textValue' ||
      child.namespaceUri !== WML_NAMESPACE_URI ||
      child.localName !== name
    )
      return child;
    found = true;
    let hasValue = false;
    const attributes = child.attributes.map((attribute) => {
      if (attribute.namespaceUri !== WML_NAMESPACE_URI || attribute.localName !== 'val')
        return attribute;
      hasValue = true;
      return attribute.value === value ? attribute : { ...attribute, value };
    });
    if (!hasValue)
      attributes.push({
        kind: 'genericExtension' as const,
        namespaceUri: WML_NAMESPACE_URI,
        localName: 'val',
        prefix: 'w',
        value,
      });
    return { ...child, attributes } as OoxmlNode;
  });
  if (!found) {
    const index =
      name === 'ilvl'
        ? 0
        : children.findIndex(
            (child) =>
              child.kind !== 'textValue' &&
              child.namespaceUri === WML_NAMESPACE_URI &&
              child.localName === 'ilvl'
          ) + 1;
    children.splice(
      index,
      0,
      build(
        nextId(),
        'generic',
        name,
        [
          {
            kind: 'genericExtension' as const,
            namespaceUri: WML_NAMESPACE_URI,
            localName: 'val',
            prefix: 'w',
            value,
          },
        ],
        []
      )
    );
  }
  return { ...node, children } as OoxmlNode;
}
