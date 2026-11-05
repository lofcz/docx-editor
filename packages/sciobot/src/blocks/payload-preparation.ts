import { replaceNode, withPart, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { activeMdxRegistry } from '../registry/active';
import { compactPreparedFields } from './payload-fields';
import { readPayloadStore } from './payload-opc-codec';
import { buildBlockIndex } from './block-index';
import { GALLERY_IMAGE_RECORD } from './gallery-record-shape';

/** Remove redundant plain payload fields only from validated DOCX input before room creation. */
export function withPreparedPayloadFields(pkg: OoxmlPackage): OoxmlPackage {
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') return pkg;
  const part = pkg.parts.get(store.part.partName)!;
  const registry = activeMdxRegistry();
  const prefix = registry.storage.nodePrefix;
  const index = buildBlockIndex(pkg);
  const galleryImages = new Set(
    [...index.byId.values()]
      .filter((block) => {
        if (
          ![2, 3].includes(block.payload.version) ||
          block.payload.type !== GALLERY_IMAGE_RECORD ||
          block.parentId !== block.payload.galleryId
        )
          return false;
        const parent = block.parentId && index.byId.get(block.parentId)?.payload;
        return (
          parent && parent.version === block.payload.version && parent.type === 'image_gallery'
        );
      })
      .map((block) => block.id)
  );
  let next = part;
  for (const record of part.root.children) {
    if (
      record.kind === 'textValue' ||
      record.localName !== 'node' ||
      record.namespaceUri !== part.root.namespaceUri
    )
      continue;
    const id = record.attributes.find(
      (attribute) => attribute.namespaceUri === '' && attribute.localName === 'id'
    )?.value;
    if (!id?.startsWith(prefix)) continue;
    const decoded = store.entries.get(id.slice(prefix.length))?.decoded;
    if (decoded?.status !== 'supported') continue;
    const definition = registry.components.get(decoded.value.type);
    if (
      !galleryImages.has(decoded.value.id) &&
      (!definition || (definition.valid && !definition.valid(decoded.value)))
    )
      continue;
    const compact = compactPreparedFields(record);
    if (compact === record) continue;
    const replaced = replaceNode(next, record.id, compact);
    if (!replaced.ok) throw new Error('Cannot prepare payload field history.');
    next = replaced.part;
  }
  return next === part ? pkg : withPart(pkg, next);
}
