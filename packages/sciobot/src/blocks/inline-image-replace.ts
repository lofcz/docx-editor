import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  lockForbidsRemoval,
  parentNodeOf,
  type ImageDecodePort,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import {
  resolveInlineImageResources,
  type InlineImageResolver,
} from '../native/inline-image-resources';
import { replaceManagedImageResource } from '../native/replace-image-resource';
import { captureReviewGroupHistory, reviewGroupHistoryRefusal } from '../review/group-history';
import { groupRemovalBudget } from '../review/group-removal-budget';
import { groupReviewFieldBudget } from '../review/group-field-budget';
import { ancestorContentWriteRefusal } from './ancestor-content-write';
import { decodeInlineImagePayload } from './inline-image-codec';
import { readManagedImage } from './inline-image-update';
import { encodeBlockPayload } from './payload-schema';

export interface InlineImageReplacement {
  readonly drawingNodeId: string;
  readonly expectedRevision: number;
  readonly reference: string;
  readonly resolve: InlineImageResolver;
  readonly decodePort: ImageDecodePort;
  readonly signal?: AbortSignal;
}
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Replace one explicitly chosen image. Native geometry, description, and identity stay intact. */
export async function replaceInlineImage(
  editor: DocxEditorInstance,
  request: InlineImageReplacement
): Promise<{ readonly success: true; readonly revision: number } | NativeEditFailure> {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Image replacement requires edit mode.');
  const session = surface.session;
  const revision = session.packageRevision();
  if (revision !== request.expectedRevision)
    return fail('stale-revision', 'Read the current image before replacing it.');
  const collaboration = surface.collaborationSession();
  const historyRefusal = reviewGroupHistoryRefusal(collaboration);
  if (historyRefusal) return fail('unsupported-operation', historyRefusal);
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
  let paragraphId: string | null = null;
  for (
    let parent = parentNodeOf(image.part, image.control.id);
    parent;
    parent = parentNodeOf(image.part, parent.id)
  ) {
    if (parent.kind === 'paragraph' && !paragraphId) paragraphId = parent.id;
    if (
      parent.kind === 'contentControl' &&
      lockForbidsRemoval(contentControlPropertiesOf(parent).lock)
    )
      return fail(
        'unsupported-operation',
        'An enclosing native control protects the image wrapper.'
      );
  }
  if (!paragraphId) return fail('unsupported-operation', 'The image paragraph is missing.');
  try {
    if (
      !decodeInlineImagePayload(
        encodeBlockPayload({
          ...image.payload,
          attributes: { ...image.payload.attributes, reference: request.reference },
        })
      )
    )
      return fail('invalid-input', 'The image reference is invalid.');
    const resources = await resolveInlineImageResources(
      [request.reference],
      request.resolve,
      request.decodePort,
      request.signal
    );
    if (
      request.signal?.aborted ||
      editor.surface !== surface ||
      surface.session !== session ||
      session.packageRevision() !== revision ||
      surface.editingMode() !== 'edit'
    )
      return fail('stale-revision', 'The document changed while the image loaded.');
    const edit = {
      blockId: image.block.id,
      partName: image.part.name,
      controlId: image.control.id,
      drawingNodeId: image.drawing.id,
      paragraphId,
      before: image.payload,
      resource: {
        reference: request.reference,
        alt: image.payload.attributes.alt,
        title: image.payload.attributes.title,
      },
    };
    const resource = resources.get(request.reference)!;
    const transform = (working: OoxmlPackage) =>
      replaceManagedImageResource(working, edit, resource);
    const prepared = transform(pkg);
    const result = captureReviewGroupHistory(
      collaboration,
      () => surface.applyAutomationOps(() => [], { kind: 'body' }, [transform], false, true),
      groupRemovalBudget(pkg, prepared),
      groupReviewFieldBudget(pkg, prepared)
    );
    return result.rejected || !result.committed
      ? fail('apply-failed', String(result.reason ?? 'The image replacement did not commit.'))
      : { success: true, revision: session.packageRevision() };
  } catch (cause) {
    return fail('apply-failed', cause instanceof Error ? cause.message : String(cause));
  }
}
