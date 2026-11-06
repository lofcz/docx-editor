import { replaceNode, type EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { MAX_XML_DEPTH, XML_DEFAULT_MAX_ELEMENTS } from '../package/xml-reader.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';
import { children, isRunOriginNamespace } from './run-origin-metadata.ts';

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

/** Pure export cleanup. Never publish primitives while constructing an immutable read. */
export function rootWithoutProvenanceNamespaces(
  root: OoxmlElement,
  witnessed: ReadonlySet<string>,
  namespaceUri?: string
): OoxmlElement {
  const unused = new Set(
    root.namespaceBindings
      .filter(
        (binding) =>
          (namespaceUri
            ? binding.namespaceUri === namespaceUri
            : isRunOriginNamespace(binding.namespaceUri)) &&
          witnessed.has(binding.prefix) &&
          binding.prefix.length <= 128
      )
      .map((binding) => binding.prefix)
  );
  if (!unused.size) return root;
  const pending: { node: OoxmlNode; depth: number }[] = [{ node: root, depth: 0 }];
  let visited = 0;
  while (pending.length && unused.size) {
    const { node, depth } = pending.pop()!;
    if (++visited > XML_DEFAULT_MAX_ELEMENTS * 2 + 1 || depth > MAX_XML_DEPTH + 1)
      throw new PackageEditRefusal(
        'Run-origin namespace traversal exceeds its limit.',
        'resource-limit'
      );
    if (node.kind !== 'textValue') {
      if (node.prefix) unused.delete(node.prefix);
      for (const attr of node.attributes) {
        if (attr.prefix) unused.delete(attr.prefix);
        for (const prefix of unused) {
          // QName-valued extension data and MC prefix lists remain opaque and intact.
          if (
            attr.value.includes(`${prefix}:`) ||
            (attr.namespaceUri === 'http://schemas.openxmlformats.org/markup-compatibility/2006' &&
              attr.value.split(/\s+/).includes(prefix))
          )
            unused.delete(prefix);
        }
      }
    }
    const nested = children(node);
    if (pending.length + nested.length > XML_DEFAULT_MAX_ELEMENTS * 2 + 1)
      throw new PackageEditRefusal(
        'Run-origin namespace traversal exceeds its limit.',
        'resource-limit'
      );
    for (let at = nested.length - 1; at >= 0; at--)
      pending.push({ node: nested[at]!, depth: depth + 1 });
  }
  if (!unused.size) return root;
  return {
    ...root,
    namespaceBindings: root.namespaceBindings.filter(
      (binding) =>
        !(namespaceUri
          ? binding.namespaceUri === namespaceUri
          : isRunOriginNamespace(binding.namespaceUri)) || !unused.has(binding.prefix)
    ),
  } as OoxmlElement;
}
