import { replaceNode, type EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';
import { rootWithoutProvenanceNamespaces } from './run-origin-export-namespaces.ts';

/** Remove only witnessed provenance aliases that serialization hoisted onto the root. */
export function pruneRunOriginNamespaces(
  part: OoxmlPart,
  witnessed: ReadonlySet<string>,
  options?: EditOptions,
  namespaceUri?: string
): OoxmlPart {
  const root = rootWithoutProvenanceNamespaces(part.root, witnessed, namespaceUri);
  if (root === part.root) return part;
  const made = replaceNode(part, part.root.id, root, options);
  if (!made.ok)
    throw new PackageEditRefusal('Run-origin namespace cleanup failed.', 'tree-invariant');
  return made.part;
}
