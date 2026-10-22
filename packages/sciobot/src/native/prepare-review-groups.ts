import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { withInitializedReviewGroupStore } from '../review/group-record-store';
import { initializedPackageBytes } from './prepared-parts';

/** Group-only initialization for hosts that prepare other application infrastructure separately. */
export function prepareReviewGroupDocument(
  document: Uint8Array
):
  | { readonly success: true; readonly document: Uint8Array; readonly changed: boolean }
  | NativeEditFailure {
  if (!ArrayBuffer.isView(document) || document.BYTES_PER_ELEMENT !== 1)
    return {
      success: false,
      code: 'invalid-input',
      error: 'Native document input must be DOCX bytes.',
    };
  const read = readOoxmlPackage(document);
  if (!read.ok)
    return {
      success: false,
      code: 'invalid-document',
      error: `Cannot open the DOCX: ${read.reason}.`,
    };
  const prepared = withInitializedReviewGroupStore(read.package);
  if (!prepared.ok) return { success: false, code: 'invalid-document', error: prepared.reason };
  try {
    return {
      success: true,
      changed: prepared.pkg !== read.package,
      document: initializedPackageBytes(document, read.package, prepared.pkg),
    };
  } catch (cause) {
    return {
      success: false,
      code: 'invalid-document',
      error: cause instanceof Error ? cause.message : String(cause),
    };
  }
}
