import {
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
  type OoxmlNode,
} from '../package/ooxml-tree.ts';
import type { WmlFreshNamespaceContext } from '../package/wml-namespace.ts';
import { wmlChildNamed } from './tree-op-table-shared.ts';

/** Patch only operation-owned WML attributes. Keep imported IDs, extensions, and children. */
export function retainTablePropertyAttributes(
  parent: OoxmlElement | undefined,
  replacement: OoxmlElement
): OoxmlElement {
  const old = parent && wmlChildNamed(parent, replacement.localName);
  if (!old) return replacement;
  const pending = new Map(
    replacement.attributes.map((attribute) => [attribute.localName, attribute])
  );
  let changed = false;
  const attributes: OoxmlAttribute[] = old.attributes.map((attribute) => {
    if (attribute.namespaceUri !== WML_NAMESPACE_URI) return attribute;
    const next = pending.get(attribute.localName);
    if (!next) return attribute;
    pending.delete(attribute.localName);
    if (attribute.value === next.value) return attribute;
    changed = true;
    return Object.freeze({ ...attribute, value: next.value });
  });
  for (const attribute of pending.values()) {
    changed = true;
    attributes.push(attribute);
  }
  return changed ? (Object.freeze({ ...old, attributes }) as OoxmlElement) : old;
}

export function freshWmlElement(
  localName: string,
  nextId: () => string,
  wml: WmlFreshNamespaceContext,
  attributes: readonly OoxmlAttribute[],
  children: readonly OoxmlNode[] = []
): OoxmlElement {
  return {
    id: nextId(),
    kind: 'generic',
    namespaceUri: WML_NAMESPACE_URI,
    localName,
    ...(wml.elementPrefix === undefined ? {} : { prefix: wml.elementPrefix }),
    namespaceBindings: [],
    attributes,
    children,
  } as OoxmlElement;
}
