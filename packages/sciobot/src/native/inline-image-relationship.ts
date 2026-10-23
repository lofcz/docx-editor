import {
  normalizePartName,
  relsPartNameFor,
  resolveImageRelationship,
  replaceNode,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

/** Keep the public image mint's relationship bound to its actual media outside /word stories. */
export function bindInlineImageRelationship(
  pkg: OoxmlPackage,
  owner: string,
  relationshipId: string,
  mediaPartName: string
): OoxmlPackage {
  const media = normalizePartName(mediaPartName);
  if (!media.ok) throw Error('The embedded media part name is invalid.');
  const owned = pkg.relationships.get(owner) ?? [];
  const resolved = resolveImageRelationship(owned, owner, relationshipId);
  if (resolved.mode === 'internal' && resolved.partName === media.partName) return pkg;
  const part = pkg.parts.get(relsPartNameFor(owner));
  if (!part) throw Error('The image owner has no native relationship part.');
  const nodes = part.root.children.filter(
    (node) =>
      node.kind !== 'textValue' &&
      node.localName === 'Relationship' &&
      node.attributes.some(
        (attr) =>
          attr.namespaceUri === '' && attr.localName === 'Id' && attr.value === relationshipId
      )
  );
  const record = owned.find((record) => record.id === relationshipId);
  if (nodes.length !== 1 || !record || record.targetMode !== 'Internal')
    throw Error('The image relationship must identify one internal record.');
  const node = nodes[0]!;
  if (node.kind === 'textValue') throw Error('The native image relationship is not an element.');
  const changed = replaceNode(part, node.id, {
    ...node,
    attributes: node.attributes.map((attr) =>
      attr.namespaceUri === '' && attr.localName === 'Target'
        ? { ...attr, value: media.partName }
        : attr
    ),
  });
  if (!changed.ok) throw Error('The native image relationship target could not change.');
  const relationships = new Map(pkg.relationships);
  relationships.set(
    owner,
    owned.map((entry) =>
      entry.id === relationshipId ? { ...entry, rawTarget: media.partName } : entry
    )
  );
  return Object.freeze({ ...withPart(pkg, changed.part), relationships });
}
