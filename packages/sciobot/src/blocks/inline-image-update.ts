import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  findNode,
  parentNodeOf,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { buildBlockIndex } from './block-index';
import { decodeInlineImagePayload } from './inline-image-codec';
import { encodeBlockPayload } from './payload-schema';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';

/** Resolve one managed drawing without treating arbitrary locked Word controls as app images. */
export function readManagedImage(pkg: OoxmlPackage, partName: string, drawingNodeId: string) {
  const part = pkg.parts.get(partName);
  const drawing = part && findNode(part, drawingNodeId);
  if (!part || drawing?.kind !== 'drawing') return null;
  const index = buildBlockIndex(pkg);
  for (
    let parent = parentNodeOf(part, drawing.id);
    parent;
    parent = parentNodeOf(part, parent.id)
  ) {
    if (parent.kind !== 'contentControl') continue;
    const block = [...index.byId.values()].find(
      (entry) => entry.controlId === parent.id && entry.partName === partName
    );
    const payload = block && decodeInlineImagePayload(encodeBlockPayload(block.payload));
    if (!block || !payload) return null;
    const drawings: OoxmlNode[] = [];
    const visit = (node: OoxmlNode) => {
      if (node.kind === 'drawing') drawings.push(node);
      else if (node.kind !== 'textValue') node.children.forEach(visit);
    };
    visit(parent);
    if (drawings.length !== 1 || drawings[0]!.id !== drawing.id) return null;
    return { block, payload, control: parent, drawing, part };
  }
  return null;
}

export interface InlineImageUpdate {
  readonly drawingNodeId: string;
  readonly expectedRevision: number;
  readonly change:
    | { readonly type: 'size'; readonly widthPoints: number; readonly heightPoints: number }
    | { readonly type: 'alt'; readonly alt: string };
}
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Update owned image metadata and native geometry in one canonical history unit. */
export function updateInlineImage(
  editor: DocxEditorInstance,
  request: InlineImageUpdate
): { success: true; revision: number } | NativeEditFailure {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Image changes require edit mode.');
  const session = surface.session;
  if (session.packageRevision() !== request.expectedRevision)
    return fail('stale-revision', 'Read the current image before changing it.');
  const pkg = session.currentPackage();
  const image = readManagedImage(pkg, session.part().name, request.drawingNodeId);
  if (!image) return fail('invalid-input', 'The managed image is missing or ambiguous.');
  const properties = contentControlPropertiesOf(image.control);
  const ancestor = ancestorContentWriteRefusal(image.part, image.control.id);
  if (
    ancestor ||
    properties.dataBinding ||
    (properties.lock && properties.lock !== 'unlocked' && properties.lock !== 'sdtContentLocked')
  )
    return fail('unsupported-operation', ancestor ?? 'The image control is protected.');
  const change = request.change;
  const attributes = { ...image.payload.attributes };
  const ops: TreeDocOp[] = [];
  if (change.type === 'size') {
    if (
      ![change.widthPoints, change.heightPoints].every(
        (value) => Number.isFinite(value) && value > 0 && value <= 158400
      )
    )
      return fail('invalid-input', 'The image dimensions are invalid.');
    attributes.widthPoints = change.widthPoints;
    attributes.heightPoints = change.heightPoints;
    ops.push({
      op: 'resizeDrawing',
      drawingNodeId: image.drawing.id,
      extentEmu: {
        cx: Math.round(change.widthPoints * 12700),
        cy: Math.round(change.heightPoints * 12700),
      },
    });
  } else {
    attributes.alt = change.alt;
    const metadata: OoxmlNode[] = [];
    const visit = (node: OoxmlNode) => {
      if (node.kind === 'drawingDocPr') metadata.push(node);
      else if (node.kind !== 'textValue') node.children.forEach(visit);
    };
    visit(image.drawing);
    if (metadata.length !== 1 || metadata[0]!.kind === 'textValue')
      return fail('unsupported-operation', 'The native picture metadata is ambiguous.');
    const title =
      metadata[0]!.attributes.find(
        (attribute) => attribute.namespaceUri === '' && attribute.localName === 'title'
      )?.value ?? '';
    ops.push({
      op: 'setDrawingMetadata',
      drawingNodeId: image.drawing.id,
      title,
      description: change.alt,
    });
  }
  const payload = { ...image.payload, attributes };
  if (!decodeInlineImagePayload(encodeBlockPayload(payload)))
    return fail('invalid-input', 'The image metadata is invalid.');
  const transform = (working: OoxmlPackage): OoxmlPackage => {
    let part = working.parts.get(image.part.name)!;
    const apply = (op: TreeDocOp) => {
      const result = applyTreeOp(part, op);
      if (!result.ok) throw Error(String(result.reason));
      part = result.part;
    };
    if (properties.lock === 'sdtContentLocked')
      apply({ op: 'setContentControlProperties', controlId: image.control.id, lock: 'unlocked' });
    ops.forEach(apply);
    if (properties.lock === 'sdtContentLocked')
      apply({
        op: 'setContentControlProperties',
        controlId: image.control.id,
        lock: 'sdtContentLocked',
      });
    const store = readPayloadStore(working);
    const entry = store.status === 'ready' && store.entries.get(image.block.id);
    if (!entry) throw Error('The image metadata disappeared.');
    const stored = withPayloadRecord(withPart(working, part), payload, entry.label);
    if (!stored.ok) throw Error(stored.reason);
    return stored.pkg;
  };
  try {
    transform(pkg);
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [transform], false, true);
    return result.rejected || !result.committed
      ? fail('apply-failed', String(result.reason ?? 'The image change did not commit.'))
      : { success: true, revision: session.packageRevision() };
  } catch (cause) {
    return fail('apply-failed', cause instanceof Error ? cause.message : String(cause));
  }
}
