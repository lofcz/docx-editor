import { withInitializedPayloadStore } from '../blocks/payload-opc-codec';
import { withAtomicMathFallbacks } from '../blocks/math-ownership';
import { withPreparedPayloadFields } from '../blocks/payload-preparation';
import { withGalleryImageRecords } from '../blocks/gallery-record-upgrade';
import { withNativeNumberingInfrastructure } from './numbering-infrastructure';
import { withTaskCompletionStyles } from './task-completion-style';
import { withPreparedTaskCompletion } from './prepare-task-completion';
import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { withInitializedReviewGroupStore } from '../review/group-record-store';
import { initializedPackageBytes } from './prepared-parts';
import {
  inlineCodeInitializationStyles,
  inlineCodeStyleState,
  withInlineCodeStyle,
} from './inline-code';

export type PreparedNativeDocument =
  | {
      readonly success: true;
      readonly document: Uint8Array;
      readonly changed: boolean;
      /** A conflicting imported style remains untouched; code formatting can report its conflict. */
      readonly codeStyleReady: boolean;
    }
  | NativeEditFailure;

/**
 * Initialize application-owned code styles, numbering infrastructure, payload storage, and review storage before opening or sharing DOCX bytes.
 *
 * Call this once before a collaboration room bootstraps, not after editors join. Otherwise two
 * offline first writers can each create the styles part, and undoing one part's creation leaves
 * the other writer's relationship dangling. Prepared documents share that infrastructure as
 * initial content; each actor's formatting remains an ordinary independent undoable edit. Derived
 * definitions for imported character styles are also initial content, so no actor owns their lifetime.
 *
 * Supported equation controls gain durable fallback ownership before room creation.
 * Valid application tasks gain their completion appearance before room creation.
 * Only changed parts are serialized. Unrelated parts keep their original bytes.
 * This does not invoke the Word export cleanup or reconstruct content from Markdown.
 */
export function prepareNativeDocument(document: Uint8Array): PreparedNativeDocument {
  if (!ArrayBuffer.isView(document) || document.BYTES_PER_ELEMENT !== 1)
    return {
      success: false,
      code: 'invalid-input',
      error: 'Native document input must be DOCX bytes.',
    };
  const loaded = readOoxmlPackage(document);
  if (!loaded.ok)
    return {
      success: false,
      code: 'invalid-document',
      error: `Cannot open the DOCX: ${loaded.reason}.`,
    };
  const original = loaded.package;
  // Group initialization also creates a missing owner relationship part. Styles need that part.
  const groups = withInitializedReviewGroupStore(original);
  if (!groups.ok) return { success: false, code: 'invalid-document', error: groups.reason };
  const payloads = withInitializedPayloadStore(groups.pkg);
  if (!payloads.ok) return { success: false, code: 'invalid-document', error: payloads.reason };
  const galleries = withGalleryImageRecords(payloads.pkg);
  if (!galleries.ok) return { success: false, code: 'invalid-document', error: galleries.reason };
  const state = inlineCodeStyleState(galleries.pkg);
  const styled =
    state === 'conflict'
      ? galleries.pkg
      : withInlineCodeStyle(galleries.pkg, inlineCodeInitializationStyles(galleries.pkg));
  if (!styled)
    return {
      success: false,
      code: 'invalid-document',
      error: 'Cannot initialize the native code style.',
    };
  let completed;
  try {
    completed = withPreparedTaskCompletion(withTaskCompletionStyles(styled));
  } catch (error) {
    return {
      success: false,
      code: 'invalid-document',
      error: error instanceof Error ? error.message : 'Cannot initialize task completion styles.',
    };
  }
  const prepared = withNativeNumberingInfrastructure(completed),
    codeStyleReady = state !== 'conflict';
  if (!prepared)
    return {
      success: false,
      code: 'invalid-document',
      error: 'Cannot initialize native numbering infrastructure.',
    };
  try {
    const owned = withAtomicMathFallbacks(withPreparedPayloadFields(prepared));
    return {
      success: true,
      document: initializedPackageBytes(document, original, owned),
      changed: owned !== original,
      codeStyleReady,
    };
  } catch (error) {
    return {
      success: false,
      code: 'invalid-document',
      error: error instanceof Error ? error.message : 'Cannot initialize the native DOCX.',
    };
  }
}
