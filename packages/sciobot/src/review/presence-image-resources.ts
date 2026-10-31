import type { BlockPayload } from '../blocks/payload-schema';
import { activeMdxRegistry } from '../registry/active';
import {
  cleanupOrphanImageMedia,
  findNode,
  resolveImageRelationship,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import type { BlockPresenceChange } from './block-presence-schema';

interface Resource {
  readonly owner: string;
  readonly media: string;
  readonly relationship: string;
}
const relationshipsIn = (node: OoxmlNode): string[] =>
  node.kind === 'textValue'
    ? []
    : [
        ...node.attributes
          .filter(
            (attribute) =>
              attribute.namespaceUri ===
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships' &&
              attribute.localName === 'embed'
          )
          .map((attribute) => attribute.value),
        ...node.children.flatMap(relationshipsIn),
      ];

/** Capture exact managed media ownership before native decisions remove its controls. */
export function presenceImageResources(
  pkg: OoxmlPackage,
  block: BlockPresenceChange
): readonly Resource[] {
  const registry = activeMdxRegistry();
  const owns = (payload: BlockPayload) => registry.components.get(payload.type)?.mediaPart;
  const members = block.members.filter((member) => owns(member.payload));
  if (!members.length) return [];
  const index = buildBlockIndex(pkg);
  return members.map((member) => {
    const bound = index.byId.get(member.payload.id);
    const part = bound && pkg.parts.get(bound.partName);
    const control = part && bound && findNode(part, bound.controlId);
    const ids = control ? relationshipsIn(control) : [];
    const media = owns(member.payload)?.(member.payload);
    const related =
      part &&
      ids.length === 1 &&
      resolveImageRelationship(pkg.relationships.get(part.name), part.name, ids[0]!);
    if (
      !part ||
      typeof media !== 'string' ||
      !related ||
      related.mode !== 'internal' ||
      related.partName !== media
    )
      throw Error('The presence image resource no longer matches its provenance.');
    return { owner: part.name, media, relationship: ids[0]! };
  });
}

/** The native collector keeps media still referenced by another live drawing. */
export function retirePresenceImageResources(
  pkg: OoxmlPackage,
  resources: readonly Resource[]
): OoxmlPackage {
  let next = pkg;
  for (const resource of resources)
    next = cleanupOrphanImageMedia(next, resource.owner, resource.media, resource.relationship);
  return next;
}
