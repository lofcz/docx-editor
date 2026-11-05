import {
  applyTreeOp,
  contentControlPropertiesOf,
  findNode,
  replaceNode,
  withCustomXmlNode,
  withPart,
  type OoxmlPackage,
  WML_NAMESPACE_URI,
} from '@docx-editor.dev/core/store';
import { activeMdxRegistry } from '../registry/active';
import { buildBlockIndex } from './block-index';
import { nativeCodeSource } from './code-native-source';
import { codeRecordShape } from './code-record';
import { readPayloadStore } from './payload-opc-codec';
import { encodeBlockPayload } from './payload-schema';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { markCodeSourceLock } from './code-source-lock';

/** Upgrade validated code before bootstrap, or only code created in the current insertion. */
export function withNativeCodeSources(
  pkg: OoxmlPackage,
  onlyIds?: ReadonlySet<string>
): OoxmlPackage {
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') return pkg;
  const storage = activeMdxRegistry().storage;
  let next = pkg;
  for (const block of buildBlockIndex(pkg).byId.values()) {
    if (
      (onlyIds && !onlyIds.has(block.id)) ||
      block.payload.version !== 1 ||
      !codeRecordShape(block.payload)
    )
      continue;
    let part = next.parts.get(block.partName)!;
    const control = findNode(part, block.controlId);
    if (!control) continue;
    const properties = contentControlPropertiesOf(control);
    const source = nativeCodeSource(part, control);
    // Imported fields, line-ending differences and modified Word fallbacks remain lossless and opaque.
    if (
      properties.dataBinding ||
      (properties.lock && properties.lock !== 'unlocked') ||
      ancestorContentWriteRefusal(part, control.id) ||
      !source ||
      source.value !== block.payload.attributes.value
    )
      continue;
    const locked = applyTreeOp(part, {
      op: 'setContentControlProperties',
      controlId: control.id,
      lock: 'contentLocked',
    });
    if (!locked.ok) throw Error('Cannot protect native code source.');
    part = locked.part;
    const lockedControl = findNode(part, control.id);
    const lock =
      lockedControl?.kind === 'contentControl' &&
      lockedControl.children
        .find((child) => child.kind === 'contentControlProperties')
        ?.children.find(
          (child) => child.namespaceUri === WML_NAMESPACE_URI && child.localName === 'lock'
        );
    if (!lock) throw Error('The native code lock is missing.');
    const marked = replaceNode(part, lock.id, markCodeSourceLock(lock));
    if (!marked.ok) throw Error('Cannot retain native code lock ownership.');
    next = withPart(next, marked.part);
    next = withCustomXmlNode(next, store.part.partName, {
      id: storage.nodePrefix + block.id,
      label: store.entries.get(block.id)!.label,
      data: encodeBlockPayload({ ...block.payload, version: 2 }),
    });
  }
  return next;
}
