import { listStateNamespace } from '../registry/adapter-names';
import {
  createNodeIdAllocator,
  findNode,
  propertyContainer,
  replaceChildren,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const owned = (node: OoxmlNode) =>
  node.kind !== 'textValue' &&
  node.namespaceUri === listStateNamespace() &&
  node.localName === 'state';

/** Dormant source check state does not turn a native bullet into a visible Word checkbox. */
export function readDormantListState(
  properties: OoxmlNode | null | undefined
): { ok: true; node?: OoxmlNode; checked?: boolean } | { ok: false; reason: string } {
  const nodes =
    properties && properties.kind !== 'textValue' ? properties.children.filter(owned) : [];
  if (!nodes.length) return { ok: true };
  const node = nodes[0]!;
  if (
    nodes.length !== 1 ||
    node.kind === 'textValue' ||
    node.children.length ||
    node.attributes.filter((attribute) => attribute.namespaceUri !== MC).length !== 1 ||
    !node.attributes.some(
      (attribute) =>
        attribute.namespaceUri === '' &&
        attribute.localName === 'checked' &&
        ['true', 'false'].includes(attribute.value)
    ) ||
    node.attributes.some(
      (attribute) =>
        attribute.namespaceUri === MC &&
        !['Ignorable', 'PreserveElements'].includes(attribute.localName)
    )
  )
    return { ok: false, reason: 'The dormant list state is invalid or unsupported.' };
  return {
    ok: true,
    node,
    checked:
      node.attributes.find(
        (attribute) => attribute.namespaceUri === '' && attribute.localName === 'checked'
      )!.value === 'true',
  };
}
export function writeDormantListState(
  part: OoxmlPart,
  paragraphId: string,
  checked: boolean | undefined
): OoxmlPart {
  const paragraph = findNode(part, paragraphId);
  if (paragraph?.kind !== 'paragraph') throw new Error('The list paragraph is missing.');
  const properties = propertyContainer(paragraph, 'paragraphProperties', 'pPr');
  const current = readDormantListState(properties);
  if (!current.ok) throw new Error(current.reason);
  if (current.checked === checked) return part;
  if (!properties || properties.kind === 'textValue')
    throw new Error('The list property holder is missing.');
  const node: OoxmlElement | undefined =
    checked === undefined
      ? undefined
      : {
          id: current.node?.id ?? createNodeIdAllocator(part)(),
          kind: 'generic',
          namespaceUri: listStateNamespace(),
          localName: 'state',
          prefix: 'scioList',
          namespaceBindings: [
            { prefix: 'scioList', namespaceUri: listStateNamespace() },
            { prefix: 'mc', namespaceUri: MC },
          ],
          attributes: [
            {
              kind: 'genericExtension',
              namespaceUri: '',
              localName: 'checked',
              value: String(checked),
            },
            {
              kind: 'genericExtension',
              namespaceUri: MC,
              prefix: 'mc',
              localName: 'Ignorable',
              value: 'scioList',
            },
            {
              kind: 'genericExtension',
              namespaceUri: MC,
              prefix: 'mc',
              localName: 'PreserveElements',
              value: 'scioList:state',
            },
          ],
          children: [],
        };
  const children = properties.children.flatMap((child) =>
    owned(child) ? (node ? [node] : []) : [child]
  );
  if (node && !current.node) children.push(node);
  const result = replaceChildren(part, properties.id, children);
  if (!result.ok) throw new Error('Dormant list state failed native validation.');
  return result.part;
}
