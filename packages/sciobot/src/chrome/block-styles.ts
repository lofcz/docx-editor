import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import { withBlockTypeStyles } from '../native/block-control-styles';
import { prepareNativeDocument } from '../native/prepare-document';
import { initializedPackageBytes } from '../native/prepared-parts';

export { withBlockTypeStyles } from '../native/block-control-styles';

/** Prepare shared infrastructure and block definitions before creating either editor or room. */
export function prepareBlockControlsDocument(document: Uint8Array) {
  const native = prepareNativeDocument(document);
  if (!native.success) return native;
  const loaded = readOoxmlPackage(native.document);
  if (!loaded.ok)
    return {
      success: false as const,
      code: 'invalid-document' as const,
      error: `Cannot read the prepared document: ${loaded.reason}.`,
    };
  const styled = withBlockTypeStyles(loaded.package);
  if (!styled)
    return {
      success: false as const,
      code: 'invalid-document' as const,
      error: 'The document has conflicting block style definitions.',
    };
  return {
    success: true as const,
    document: initializedPackageBytes(native.document, loaded.package, styled),
    changed: native.changed || styled !== loaded.package,
  };
}
