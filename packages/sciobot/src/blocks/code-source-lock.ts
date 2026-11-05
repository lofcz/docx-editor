import {
  contentControlPropertiesOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';

const NAMESPACE = 'urn:sciobot:code-source-lock:1';
const MC_NAMESPACE_URI = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** Only the exact lock created during preparation belongs to code authoring. */
export function ownedCodeSourceLock(control: OoxmlNode) {
  if (control.kind !== 'contentControl') return null;
  const properties = contentControlPropertiesOf(control);
  if (properties.lock !== 'contentLocked' || properties.dataBinding) return null;
  const propertyNodes = control.children.filter(
    (child) => child.kind === 'contentControlProperties'
  );
  if (propertyNodes.length !== 1) return null;
  const locks = propertyNodes[0]!.children.filter(
    (child) => child.namespaceUri === WML_NAMESPACE_URI && child.localName === 'lock'
  );
  const lock = locks[0];
  if (locks.length !== 1 || lock?.kind !== 'generic' || lock.children.length) return null;
  const marker = lock.attributes.filter(
    (attribute) => attribute.namespaceUri === NAMESPACE && attribute.localName === 'sourceOwner'
  );
  return marker.length === 1 && marker[0]!.value === '1' ? lock : null;
}

/** The caller creates this lock on an otherwise unlocked control before room bootstrap. */
export function markCodeSourceLock(lock: OoxmlNode): OoxmlNode {
  if (
    lock.kind !== 'generic' ||
    lock.namespaceUri !== WML_NAMESPACE_URI ||
    lock.localName !== 'lock' ||
    lock.children.length ||
    lock.attributes.some((attribute) => attribute.namespaceUri !== WML_NAMESPACE_URI)
  )
    throw Error('Only a newly created native code lock can receive ownership.');
  return {
    ...lock,
    namespaceBindings: [
      ...lock.namespaceBindings,
      { prefix: 'scioCode', namespaceUri: NAMESPACE },
      { prefix: 'scioCodeMc', namespaceUri: MC_NAMESPACE_URI },
    ],
    attributes: [
      ...lock.attributes,
      {
        kind: 'genericExtension',
        namespaceUri: NAMESPACE,
        prefix: 'scioCode',
        localName: 'sourceOwner',
        value: '1',
      },
      {
        kind: 'genericExtension',
        namespaceUri: MC_NAMESPACE_URI,
        prefix: 'scioCodeMc',
        localName: 'Ignorable',
        value: 'scioCode',
      },
      {
        kind: 'genericExtension',
        namespaceUri: MC_NAMESPACE_URI,
        prefix: 'scioCodeMc',
        localName: 'PreserveAttributes',
        value: 'scioCode:sourceOwner',
      },
    ],
  };
}
