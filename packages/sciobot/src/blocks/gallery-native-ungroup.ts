import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  contentControlContentNodeOf,
  findNode,
  withoutCustomXmlNode,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { activeMdxRegistry } from '../registry/active';
import { captureReviewGroupHistory, reviewGroupHistoryRefusal } from '../review/group-history';
import { groupRemovalBudget } from '../review/group-removal-budget';
import { groupReviewFieldBudget } from '../review/group-field-budget';
import type { GalleryState } from './gallery';
import { readPayloadStore } from './payload-opc-codec';

/** Unwrap native owners. Original image and caption nodes remain in their current order. */
export function ungroupNativeGallery(editor: DocxEditorInstance, gallery: GalleryState) {
  const surface = editor.surface!;
  const refusal = reviewGroupHistoryRefusal(surface.collaborationSession());
  if (refusal) throw Error(refusal);
  const pkg = surface.session.currentPackage();
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') throw Error('The gallery metadata is unavailable.');
  const transform = (working: OoxmlPackage) => {
    let part = working.parts.get(gallery.partName)!;
    const control = findNode(part, gallery.controlId);
    const content = control && contentControlContentNodeOf(control);
    if (!content || content.children.length !== gallery.images.length)
      throw Error('The gallery owner changed.');
    for (const image of gallery.images) {
      const removed = applyTreeOp(part, {
        op: 'removeContentControl',
        controlId: image.controlId,
        keepContent: true,
      });
      if (!removed.ok) throw Error(`The image cannot be ungrouped: ${removed.reason}.`);
      part = removed.part;
    }
    const removed = applyTreeOp(part, {
      op: 'removeContentControl',
      controlId: gallery.controlId,
      keepContent: true,
    });
    if (!removed.ok) throw Error(`The gallery cannot be ungrouped: ${removed.reason}.`);
    let next = withPart(working, removed.part);
    for (const id of [gallery.id, ...gallery.images.map((image) => image.id)])
      next = withoutCustomXmlNode(
        next,
        store.part.partName,
        activeMdxRegistry().storage.nodePrefix + id
      );
    if (readPayloadStore(next).status !== 'ready')
      throw Error('The ungrouped metadata is invalid.');
    return next;
  };
  const candidate = transform(pkg);
  const result = captureReviewGroupHistory(
    surface.collaborationSession(),
    () => surface.applyAutomationOps(() => [], { kind: 'body' }, [transform], false, true),
    groupRemovalBudget(pkg, candidate),
    groupReviewFieldBudget(pkg, candidate)
  );
  if (result.rejected || !result.committed)
    throw Error(String(result.reason ?? 'The gallery did not change.'));
  return { success: true as const, revision: surface.session.packageRevision() };
}
