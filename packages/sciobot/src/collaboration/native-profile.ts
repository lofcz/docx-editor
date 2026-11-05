import type * as Y from 'yjs';

/** The exact native record representation evaluated by private review ownership. */
export function hasEvaluatedNativeProfile(doc: Y.Doc): boolean {
  if (!doc.share.has('docx-package-meta-v1')) return false;
  let meta: Y.Map<unknown>;
  try {
    // Freshly decoded roots need their concrete Yjs type. An existing wrong type refuses.
    meta = doc.getMap('docx-package-meta-v1');
  } catch {
    return false;
  }
  return (
    meta.get('protocolVersion') === 1 &&
    meta.get('sharedSchemaVersion') === 10 &&
    meta.get('repairVersion') === 17 &&
    meta.get('canonicalModelVersion') === 1
  );
}
