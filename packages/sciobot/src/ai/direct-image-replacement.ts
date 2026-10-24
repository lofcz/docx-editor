import {
  applyTreeOp,
  cleanupOrphanImageMedia,
  contentControlPropertiesOf,
  findNode,
  parentNodeOf,
  resolveImageRelationship,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import { jsonDataKey } from '../diff/target-key';
import type { CanonicalEdit, CanonicalEditPlan } from '../diff/plan';
import { embedInlineImage } from '../native/inline-image-media';
import {
  preparedInlineImageResources,
  type PreparedInlineImages,
} from '../native/inline-image-resources';
import { refuseWrite } from '../diff/fragment-seeds';

type ImageEdit = Extract<CanonicalEdit, { kind: 'inline-image' }>;
/** Resource and provenance change together; the native drawing, control, layout and link IDs survive. */
export function completeImageReplacement(
  pkg: OoxmlPackage,
  edit: ImageEdit,
  plan: CanonicalEditPlan,
  handle?: PreparedInlineImages
): OoxmlPackage {
  let next = pkg;
  const block = buildBlockIndex(pkg).byId.get(edit.blockId);
  const part = pkg.parts.get(edit.partName);
  const control = part && findNode(part, edit.controlId);
  const drawing = part && findNode(part, edit.drawingNodeId);
  if (
    !block ||
    !part ||
    control?.kind !== 'contentControl' ||
    drawing?.kind !== 'drawing' ||
    block.partName !== edit.partName ||
    block.controlId !== edit.controlId ||
    jsonDataKey(block.payload) !== jsonDataKey(edit.before)
  )
    refuseWrite('The managed image no longer matches its planned source.', 'stale-revision');
  const contains = (node: OoxmlNode, id: string): boolean =>
    node.id === id ||
    (node.kind !== 'textValue' && node.children.some((child) => contains(child, id)));
  if (!contains(control, drawing.id))
    refuseWrite('The drawing left its managed control.', 'stale-revision');
  let parent = parentNodeOf(part, control.id);
  while (parent && parent.kind !== 'paragraph') parent = parentNodeOf(part, parent.id);
  if (parent?.id !== edit.paragraphId)
    refuseWrite('The managed image left its planned paragraph.', 'stale-revision');
  const embeddedReferences: string[] = [];
  const metadata: OoxmlNode[] = [];
  let linked = false;
  const scan = (node: OoxmlNode) => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'drawingDocPr') metadata.push(node);
    if (
      node.attributes.some(
        (attribute) =>
          attribute.namespaceUri ===
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships' &&
          attribute.localName === 'link'
      )
    )
      linked = true;
    for (const attribute of node.attributes)
      if (
        attribute.namespaceUri ===
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships' &&
        attribute.localName === 'embed'
      )
        embeddedReferences.push(attribute.value);
    node.children.forEach(scan);
  };
  scan(drawing);
  const previousRelationship = embeddedReferences[0];
  if (embeddedReferences.length !== 1 || !previousRelationship || linked || metadata.length !== 1)
    refuseWrite(
      'Additional picture resource branches require explicit native reconciliation.',
      'unsupported-operation'
    );
  const readMetadata = (name: string) => {
    const node = metadata[0]!;
    return node.kind === 'textValue'
      ? ''
      : (node.attributes.find(
          (attribute) => attribute.namespaceUri === '' && attribute.localName === name
        )?.value ?? '');
  };
  const previous = resolveImageRelationship(
    pkg.relationships.get(part.name),
    part.name,
    previousRelationship
  );
  if (previous.mode !== 'internal' || previous.partName !== edit.before.attributes.mediaPartName)
    refuseWrite('The native image resource no longer matches its provenance.', 'stale-revision');
  const changed = edit.resource.reference !== edit.before.attributes.reference;
  const resource = changed
    ? preparedInlineImageResources(handle, plan)?.get(edit.resource.reference)
    : undefined;
  if (changed && !resource)
    refuseWrite(
      'Image replacement requires revision-bound resource preparation.',
      'unsupported-operation'
    );
  const properties = contentControlPropertiesOf(control);
  const apply = (op: TreeDocOp) => {
    const current = next.parts.get(part.name)!;
    const result = applyTreeOp(current, op);
    if (!result.ok)
      refuseWrite(`Native image replacement refused: ${result.reason}.`, 'apply-failed');
    next = withPart(next, result.part);
  };
  // Only the application's own immutable atom lock can be temporarily released.
  // Explicit wrapper locks, bindings and ancestor protection retain native admission.
  if (properties.lock === 'sdtContentLocked')
    apply({ op: 'setContentControlProperties', controlId: control.id, lock: 'unlocked' });
  let mediaPartName = edit.before.attributes.mediaPartName;
  const previousBytes =
    next.partBytes.get(previous.partName) ?? next.partBytes.get(previous.partName.slice(1));
  const replaceBitmap =
    !!resource &&
    (!previousBytes ||
      previousBytes.length !== resource.bytes.length ||
      previousBytes.some((byte, index) => byte !== resource.bytes[index]));
  if (resource && replaceBitmap) {
    const embedded = embedInlineImage(next, part.name, crypto.randomUUID(), resource);
    next = embedded.pkg;
    mediaPartName = embedded.partName.startsWith('/') ? embedded.partName : '/' + embedded.partName;
    apply({
      op: 'replaceDrawingResource',
      drawingNodeId: drawing.id,
      relationshipId: embedded.relationshipId,
    });
  }
  if (
    edit.resource.alt !== edit.before.attributes.alt ||
    edit.resource.title !== edit.before.attributes.title
  )
    apply({
      op: 'setDrawingMetadata',
      drawingNodeId: drawing.id,
      title:
        edit.resource.title === edit.before.attributes.title
          ? readMetadata('title')
          : (edit.resource.title ?? ''),
      description:
        edit.resource.alt === edit.before.attributes.alt
          ? readMetadata('descr')
          : edit.resource.alt,
    });
  if (properties.lock === 'sdtContentLocked')
    apply({ op: 'setContentControlProperties', controlId: control.id, lock: 'sdtContentLocked' });
  const store = readPayloadStore(next);
  const entry = store.status === 'ready' && store.entries.get(edit.blockId);
  if (!entry) refuseWrite('The image provenance record disappeared.', 'stale-revision');
  const stored = withPayloadRecord(
    next,
    {
      ...edit.before,
      attributes: {
        ...edit.before.attributes,
        ...edit.resource,
        contentHash: resource?.contentHash ?? edit.before.attributes.contentHash,
        mediaPartName,
      },
    },
    entry.label
  );
  if (!stored.ok)
    refuseWrite(`Image provenance replacement refused: ${stored.reason}.`, 'apply-failed');
  next = stored.pkg;
  return replaceBitmap
    ? cleanupOrphanImageMedia(next, part.name, previous.partName, previousRelationship)
    : next;
}
