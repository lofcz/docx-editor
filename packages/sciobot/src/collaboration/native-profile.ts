import type * as Y from 'yjs';

/** The exact native record representation evaluated by private review ownership. */
export function hasEvaluatedNativeProfile(doc: Y.Doc): boolean {
  const meta = doc.share.has('docx-package-meta-v1') ? doc.getMap('docx-package-meta-v1') : null;
  return (
    meta !== null &&
    meta.get('protocolVersion') === 1 &&
    meta.get('sharedSchemaVersion') === 10 &&
    meta.get('repairVersion') === 17 &&
    meta.get('canonicalModelVersion') === 1
  );
}
