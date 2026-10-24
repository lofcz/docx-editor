import {
  allocateContentControlId,
  contentControlPropertiesNodeOf,
  contentControlPropertiesOf,
  createNodeIdAllocator,
  findNode,
  orderedContentControlProperties,
  replaceChildren,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

/** Raw paragraph fragments need the same Word control identity as native control insertions. */
export function withImageControlId(
  pkg: OoxmlPackage,
  partName: string,
  controlId: string
): OoxmlPackage {
  const part = pkg.parts.get(partName)!;
  const control = findNode(part, controlId);
  if (control?.kind !== 'contentControl') throw Error('The image control is missing.');
  if (contentControlPropertiesOf(control).id !== undefined) return pkg;
  const properties = contentControlPropertiesNodeOf(control);
  const id = allocateContentControlId(part.root);
  if (!properties || id === null) throw Error('The image Word control identity is unavailable.');
  const mint = createNodeIdAllocator(part);
  const identity: OoxmlNode = {
    id: mint(),
    kind: 'generic',
    localName: 'id',
    prefix: 'w',
    namespaceUri: WML_NAMESPACE_URI,
    namespaceBindings: [],
    attributes: [
      {
        kind: 'genericExtension',
        namespaceUri: WML_NAMESPACE_URI,
        prefix: 'w',
        localName: 'val',
        value: String(id),
      },
    ],
    children: [],
  };
  const bound = replaceChildren(
    part,
    properties.id,
    orderedContentControlProperties([...properties.children, identity])
  );
  if (!bound.ok) throw Error('The image Word control identity could not be stored.');
  return withPart(pkg, bound.part);
}
