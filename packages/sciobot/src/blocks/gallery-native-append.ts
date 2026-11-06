import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlContentNodeOf,
  contentControlPropertiesOf,
  findNode,
  insertChildren,
  lockForbidsEdit,
  withPart,
  type ImageDecodePort,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import {
  resolveInlineImageResources,
  type InlineImageResolver,
} from '../native/inline-image-resources';
import { captureReviewGroupHistory, reviewGroupHistoryRefusal } from '../review/group-history';
import { groupRemovalBudget } from '../review/group-removal-budget';
import { groupReviewFieldBudget } from '../review/group-field-budget';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { readGallery } from './gallery';
import { buildNativeGalleryImage } from './gallery-native-build';
import { withInlineControlId } from '../native/inline-image-control-id';

const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});
/** Prepare explicitly selected image files, then append them to their captured native gallery owner. */
export async function appendNativeGalleryImages(
  editor: DocxEditorInstance,
  request: {
    blockId: string;
    images: readonly { reference: string; caption?: string; alt?: string; title?: string | null }[];
    resolve: InlineImageResolver;
    decodePort: ImageDecodePort;
    expectedRevision?: number;
    signal?: AbortSignal;
  }
): Promise<{ success: true; revision: number } | NativeEditFailure> {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit' || surface.activeScope().kind !== 'body')
    return fail('unsupported-operation', 'Gallery changes require body editing.');
  const session = surface.session,
    revision = session.packageRevision(),
    pkg = session.currentPackage();
  if (request.expectedRevision !== undefined && revision !== request.expectedRevision)
    return fail('stale-revision', 'Read the current gallery before adding images.');
  const history = reviewGroupHistoryRefusal(surface.collaborationSession());
  if (history) return fail('unsupported-operation', history);
  const gallery = readGallery(pkg, request.blockId);
  if (gallery?.payload.version !== 3 || gallery.partName !== session.part().name)
    return fail('invalid-input', 'The native gallery is missing from the active story.');
  if (
    !request.images.length ||
    gallery.images.length + request.images.length > 100 ||
    request.images.some(
      (image) =>
        typeof image.reference !== 'string' ||
        !image.reference.length ||
        image.reference.length > 2048 ||
        [image.caption, image.alt, image.title].some(
          (text) =>
            text !== undefined && text !== null && (typeof text !== 'string' || text.length > 2048)
        )
    )
  )
    return fail('invalid-input', 'Add bounded image references and captions.');
  const control = findNode(session.part(), gallery.controlId)!;
  const properties = contentControlPropertiesOf(control);
  const ancestor = ancestorContentWriteRefusal(session.part(), control.id);
  if (ancestor || properties.dataBinding || lockForbidsEdit(properties.lock))
    return fail('unsupported-operation', ancestor ?? 'The gallery content is protected.');
  try {
    const resources = await resolveInlineImageResources(
      request.images.map((image) => image.reference),
      request.resolve,
      request.decodePort,
      request.signal
    );
    if (
      request.signal?.aborted ||
      editor.surface !== surface ||
      surface.session !== session ||
      session.packageRevision() !== revision ||
      surface.editingMode() !== 'edit' ||
      surface.activeScope().kind !== 'body'
    )
      return fail('stale-revision', 'The document changed while the gallery images loaded.');
    const ids = request.images.map(() => ({
      imageId: crypto.randomUUID(),
      inlineId: crypto.randomUUID(),
    }));
    const transform = (working: OoxmlPackage) => {
      let next = working;
      for (const [at, image] of request.images.entries()) {
        const built = buildNativeGalleryImage(next, gallery.partName, {
          ...image,
          ...ids[at]!,
          galleryId: gallery.id,
          resource: resources.get(image.reference)!,
        });
        const part = built.pkg.parts.get(gallery.partName)!;
        const parent = findNode(part, gallery.controlId)!;
        const content = contentControlContentNodeOf(parent)!;
        const inserted = insertChildren(part, content.id, content.children.length, [built.control]);
        if (!inserted.ok) throw Error('The native gallery image cannot be added.');
        next = withInlineControlId(
          withPart(built.pkg, inserted.part),
          gallery.partName,
          built.inlineControlId
        );
      }
      return next;
    };
    const candidate = transform(pkg);
    if (!readGallery(candidate, gallery.id))
      throw Error('The native gallery image ownership is invalid.');
    const result = captureReviewGroupHistory(
      surface.collaborationSession(),
      () => surface.applyAutomationOps(() => [], { kind: 'body' }, [transform], false, true),
      groupRemovalBudget(pkg, candidate),
      groupReviewFieldBudget(pkg, candidate)
    );
    if (result.rejected || !result.committed)
      return fail('apply-failed', String(result.reason ?? 'The gallery did not change.'));
    return { success: true, revision: session.packageRevision() };
  } catch (cause) {
    return fail('apply-failed', cause instanceof Error ? cause.message : String(cause));
  }
}
