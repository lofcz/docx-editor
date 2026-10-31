import { activeStorage } from '../registry/active';
import { type OoxmlNode, type OoxmlElement, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import type { DocumentView, NativeEditFailure } from '../contracts/document';
import { bindBlockPresenceRecord } from './block-presence-plan';
import type { BlockPresenceRecord } from './block-presence-schema';

/** Derive one read snapshot. Never call canonical write primitives while projecting a presence record. */
export function projectBlockPresence(
  pkg: OoxmlPackage,
  record: BlockPresenceRecord,
  kind: DocumentView
): { readonly success: true; readonly pkg: OoxmlPackage } | NativeEditFailure {
  if (kind !== 'committed' && kind !== 'working')
    return { success: false, code: 'invalid-input', error: 'Unknown presence view.' };
  const bound = bindBlockPresenceRecord(pkg, record);
  if (!bound.success) return bound;
  const controls = new Set<string>(),
    payloads = new Set<string>();
  for (const block of bound.record.blocks) {
    if ((kind === 'committed' ? block.before : block.after).present) continue;
    controls.add(bound.controlIds.get(block.blockId)!);
    for (const member of block.members) payloads.add(member.payload.id);
  }
  if (!controls.size) return { success: true, pkg };
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready')
    return {
      success: false,
      code: 'invalid-document',
      error: 'The presence payload store is missing.',
    };
  const visit = (node: OoxmlNode, metadata: boolean): OoxmlNode => {
    if (node.kind === 'textValue') return node;
    let changed = false;
    const children: OoxmlNode[] = [];
    for (const child of node.children) {
      const payload =
        metadata &&
        child.kind !== 'textValue' &&
        child.namespaceUri === activeStorage().namespace &&
        child.localName === 'node'
          ? child.attributes.find(
              (attribute) => attribute.namespaceUri === '' && attribute.localName === 'id'
            )?.value
          : undefined;
      if (
        controls.has(child.id) ||
        (payload?.startsWith('scio_') && payloads.has(payload.slice('scio_'.length)))
      ) {
        changed = true;
        continue;
      }
      const next = visit(child, false);
      changed ||= next !== child;
      children.push(next);
    }
    return changed ? ({ ...node, children } as OoxmlElement) : node;
  };
  const parts = new Map(pkg.parts);
  for (const name of new Set([
    ...bound.record.blocks.map((block) => block.partName),
    store.part.partName,
  ])) {
    const part = pkg.parts.get(name)!;
    const root = visit(part.root, name === store.part.partName) as OoxmlElement;
    if (root !== part.root) parts.set(name, { ...part, root });
  }
  return { success: true, pkg: Object.freeze({ ...pkg, parts }) };
}
