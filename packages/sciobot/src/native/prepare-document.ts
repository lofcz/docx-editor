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
 * Initialize application-owned code styles and review storage before opening or sharing DOCX bytes.
 *
 * Call this once before a collaboration room bootstraps, not after editors join. Otherwise two
 * offline first writers can each create the styles part, and undoing one part's creation leaves
 * the other writer's relationship dangling. Prepared documents share that infrastructure as
 * initial content; each actor's formatting remains an ordinary independent undoable edit. Derived
 * definitions for imported character styles are also initial content, so no actor owns their lifetime.
 *
 * Only changed infrastructure/relationship parts and content types are serialized. Original body,
 * furniture, media, pending revisions, and recovery metadata keep their exact input bytes.
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
  const state = inlineCodeStyleState(groups.pkg);
  const styled =
    state === 'conflict'
      ? groups.pkg
      : withInlineCodeStyle(groups.pkg, inlineCodeInitializationStyles(groups.pkg));
  if (!styled)
    return {
      success: false,
      code: 'invalid-document',
      error: 'Cannot initialize the native code style.',
    };
  const prepared = styled,
    codeStyleReady = state !== 'conflict';
  try {
    return {
      success: true,
      document: initializedPackageBytes(document, original, prepared),
      changed: prepared !== original,
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
