/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { ContentTypeIndex, OoxmlPart } from '@docx-editor.dev/core/store';
import type { EncodedRelationship, RepairIssue } from './schema.ts';
import type { DocumentRegistry } from './registry.ts';
import type { BlobBytesStore } from './seed.ts';
import type { RetiredImageResources } from './retired-image-resources.ts';
import type { PackageProjectionCache } from './materialize-package-cache.ts';
import type { MaterializeResult } from './materialize.ts';
import { typeUntypedParts } from './materialize-rescued-parts.ts';
import { withNumberingTypeBytes } from './materialize-numbering-types.ts';
interface PackageAssemblyState {
  readonly registry: DocumentRegistry;
  readonly blobs: BlobBytesStore;
  readonly packageCache: PackageProjectionCache;
  readonly retiredImages: RetiredImageResources;
  readonly customXmlRels: readonly EncodedRelationship[];
  readonly customXmlOverrides: ReadonlyMap<string, string>;
  readonly issues: readonly RepairIssue[];
}
export function assembleMaterializedPackage(
  state: PackageAssemblyState,
  parts: Map<string, OoxmlPart>
): MaterializeResult {
  const assembly = state.packageCache.relationshipAssembly(state.customXmlRels);
  if (!assembly) return { ok: false, code: 'invalid-relationships', issues: [...state.issues] };
  const partBytes = state.packageCache.resolvePartBytes(state.blobs);
  if (!partBytes) return { ok: false, code: 'missing-blob', issues: [...state.issues] };
  const overrides = new Map(state.packageCache.contentTypeOverrides());
  for (const [name, mediaType] of state.customXmlOverrides) overrides.set(name, mediaType);
  const contentTypes: ContentTypeIndex = {
    defaults: state.packageCache.contentTypeDefaults(),
    overrides,
  };
  const recovered = state.retiredImages.recover(parts, assembly.byOwner, partBytes, overrides);
  if (!recovered) return { ok: false, code: 'missing-blob', issues: [...state.issues] };
  typeUntypedParts(parts, contentTypes, overrides);
  const mainDocumentPart =
    state.registry.mainDocumentPart() ||
    [...parts.keys()].find((name) => name.endsWith('document.xml')) ||
    '';
  return {
    ok: true,
    package: withNumberingTypeBytes(
      Object.freeze({
        parts,
        partBytes: recovered.bytes,
        relationships: recovered.relationships,
        externalTargets: assembly.externalTargets,
        contentTypes,
        mainDocumentPart,
      })
    ),
    issues: [...state.issues],
  };
}
