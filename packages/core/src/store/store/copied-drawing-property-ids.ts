import type { OoxmlNode } from '../package/ooxml-tree.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import { mintFragmentUniqueIds } from './clipboard-fragment-unique-ids.ts';
import { rewriteIdentifiers } from './clipboard-fragment-identifiers.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';
import { MAX_UNSIGNED_INT } from '../package/drawing-package-edit.ts';

const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

/**
 * Give copied drawing properties independent package-wide `wp:docPr/@id` values.
 * Call this inside the owning store transaction to retain its collaboration actor allocation.
 * The source package and blocks stay unchanged. This does not insert blocks, copy resources,
 * or allocate canonical node, paragraph, bookmark, revision, or other drawing identities.
 * Same-story copies can retain their existing embedded-media relationships.
 *
 * @param pkg - The bounded canonical package that receives the copied blocks.
 * @param blocks - Canonical blocks from the bounded OOXML reader.
 * @param maxNodes - Positive node budget, capped at the native limit of 500,000.
 * @param maxDepth - Positive nesting budget, capped at the native limit of 64.
 * @throws An error for invalid budgets, ambiguous source identifiers, or exhausted allocation.
 * @public
 */
export function freshenCopiedDrawingPropertyIds(
  pkg: OoxmlPackage,
  blocks: readonly OoxmlNode[],
  maxNodes: number,
  maxDepth: number
): readonly OoxmlNode[] {
  if (!Number.isInteger(maxNodes) || maxNodes < 1 || !Number.isInteger(maxDepth) || maxDepth < 1)
    throw new PackageEditRefusal('Copied drawing budgets are invalid.', 'fragment-resource-budget');
  maxNodes = Math.min(maxNodes, 500_000);
  maxDepth = Math.min(maxDepth, 64);
  const ids = new Set<string>();
  let total = 0;
  for (const root of blocks) {
    const pending = [{ node: root, depth: 0 }];
    while (pending.length) {
      const { node, depth } = pending.pop()!;
      if (++total > maxNodes || depth > maxDepth)
        throw new PackageEditRefusal(
          'Copied drawings exceed fragment limits.',
          'fragment-resource-budget'
        );
      if (node.kind === 'textValue') continue;
      if (node.namespaceUri === WP && node.localName === 'docPr') {
        const raw = node.attributes.find(
          (attribute) => attribute.namespaceUri === '' && attribute.localName === 'id'
        )?.value;
        if (
          node.kind !== 'drawingDocPr' ||
          !raw ||
          !/^\d{1,10}$/.test(raw) ||
          Number(raw) < 1 ||
          Number(raw) > MAX_UNSIGNED_INT ||
          ids.has(String(Number(raw)))
        )
          throw new PackageEditRefusal(
            'Copied drawing property identifiers are ambiguous.',
            'fragment-invalid-block'
          );
        ids.add(String(Number(raw)));
      }
      for (const child of node.children) pending.push({ node: child, depth: depth + 1 });
    }
  }
  if (!ids.size) return blocks;
  const ownerPart = pkg.parts.get(pkg.mainDocumentPart);
  if (!ownerPart) throw new PackageEditRefusal('The target document part is missing.');
  const maps = mintFragmentUniqueIds({
    pkg,
    ownerPart,
    notesParts: [],
    travelling: blocks,
    fragmentBookmarkIds: [],
    hasRevision: false,
    hasDocPr: true,
  });
  if (!maps)
    throw new PackageEditRefusal('No independent drawing property identifier is available.');
  return blocks.map((node) => rewriteIdentifiers(node, { docPrIds: maps.docPrIds }));
}
