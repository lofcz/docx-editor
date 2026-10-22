import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlElement } from '../package/ooxml-tree.ts';
import { textWithHyphens } from './tree-op-inline-elements.ts';

export function wmlElement(
  nextId: () => string,
  localName: string,
  options: {
    readonly kind?: OoxmlNode['kind'];
    readonly attributes?: readonly (readonly [string, string])[];
    readonly children?: readonly OoxmlNode[];
  } = {}
): OoxmlElement {
  return {
    id: nextId(),
    kind: options.kind ?? 'generic',
    namespaceUri: WML_NAMESPACE_URI,
    localName,
    prefix: 'w',
    namespaceBindings: [],
    attributes: (options.attributes ?? []).map(([name, value]) => ({
      kind: 'generic',
      namespaceUri: WML_NAMESPACE_URI,
      localName: name,
      prefix: 'w',
      value,
    })),
    children: options.children ?? [],
  } as unknown as OoxmlElement;
}

export function textRun(
  nextId: () => string,
  text: string,
  properties: OoxmlNode | undefined
): OoxmlNode {
  const content = textWithHyphens(nextId, text);
  return {
    id: nextId(),
    kind: 'run',
    namespaceUri: WML_NAMESPACE_URI,
    localName: 'r',
    prefix: 'w',
    namespaceBindings: [],
    attributes: [],
    children: properties ? [properties, ...content] : content,
  } as unknown as OoxmlNode;
}
