import {
  contentControlPropertiesOf,
  findNode,
  isAtomicContentControl,
  replaceNode,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { nativeImageFields } from './native-image-fields';

/** New clipboard-owned controls receive new tags without changing their content locks. */
export function withCopiedAtomicTags(
  pkg: OoxmlPackage,
  copies: readonly {
    partName: string;
    controlId: string;
    previousTag: string;
    tag: string;
    image?: boolean;
  }[]
): OoxmlPackage {
  let next = pkg;
  for (const copy of copies) {
    const part = next.parts.get(copy.partName)!;
    const control = findNode(part, copy.controlId);
    if (
      !control ||
      (!isAtomicContentControl(control) &&
        !(
          copy.image &&
          nativeImageFields(control) &&
          contentControlPropertiesOf(control).lock === 'sdtContentLocked' &&
          !contentControlPropertiesOf(control).dataBinding
        )) ||
      contentControlPropertiesOf(control).tag !== copy.previousTag
    )
      throw Error('The copied atomic control identity changed.');
    const properties =
      control.kind !== 'textValue' &&
      control.children.find(
        (node) =>
          node.kind !== 'textValue' &&
          node.namespaceUri === WML_NAMESPACE_URI &&
          node.localName === 'sdtPr'
      );
    const tag =
      properties &&
      properties.kind !== 'textValue' &&
      properties.children.find(
        (node) =>
          node.kind !== 'textValue' &&
          node.namespaceUri === WML_NAMESPACE_URI &&
          node.localName === 'tag'
      );
    if (!tag || tag.kind === 'textValue') throw Error('The copied atomic control tag is missing.');
    const changed = replaceNode(part, tag.id, {
      ...tag,
      attributes: tag.attributes.map((attribute) =>
        attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'val'
          ? { ...attribute, value: copy.tag }
          : attribute
      ),
    } as OoxmlNode);
    if (!changed.ok) throw Error('The copied atomic control tag cannot change.');
    next = withPart(next, changed.part);
  }
  return next;
}
