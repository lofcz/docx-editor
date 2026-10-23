import { sha256FontBytes } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  contentControlLevelOf,
  contentControlPropertiesOf,
  contentControlsIn,
  findNode,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { blockIdFromTag, readPayloadStore } from '../blocks/payload-opc-codec';
import { BLOCK_TAG_PREFIX } from '../blocks/payload-schema';
import type { NativeEditFailure } from '../contracts/document';
import { BLOCK_PRESENCE_LIMITS, type BlockPresenceMember } from './block-presence-schema';

export interface PresenceOwner {
  readonly blockId: string;
  readonly controlId: string;
  readonly partName: string;
  readonly members: readonly BlockPresenceMember[];
  readonly nativeFingerprint: `sha256:${string}`;
}
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-document'
): NativeEditFailure => ({ success: false, error, code });
function fingerprint(node: OoxmlNode): string | null {
  const pending = [{ node, depth: 0 }];
  let nodes = 0,
    bytes = 0;
  while (pending.length) {
    const entry = pending.pop()!;
    if (++nodes > BLOCK_PRESENCE_LIMITS.maxNodes || entry.depth > 256) return null;
    if (entry.node.kind === 'textValue') bytes += entry.node.value.length;
    else {
      bytes += entry.node.namespaceUri.length + entry.node.localName.length;
      for (const attribute of entry.node.attributes)
        bytes +=
          attribute.localName.length + attribute.namespaceUri.length + attribute.value.length;
      for (const child of entry.node.children)
        pending.push({ node: child, depth: entry.depth + 1 });
    }
    if (bytes > BLOCK_PRESENCE_LIMITS.maxBytes) return null;
  }
  const encoded = new TextEncoder().encode(canonicalOoxmlFingerprint(node));
  return encoded.length <= BLOCK_PRESENCE_LIMITS.maxBytes ? sha256FontBytes(encoded) : null;
}
/** Bind stable block IDs to this load's exact controls and bounded semantic ownership proofs. */
export function readPresenceOwners(
  pkg: OoxmlPackage,
  roots: readonly string[]
):
  | { readonly success: true; readonly owners: ReadonlyMap<string, PresenceOwner> }
  | NativeEditFailure {
  const index = buildBlockIndex(pkg),
    payloads = readPayloadStore(pkg);
  if (payloads.status !== 'ready')
    return fail('Block presence needs a valid application payload store.');
  const owners = new Map<string, PresenceOwner>(),
    allMembers = new Set<string>();
  for (const blockId of roots) {
    if (owners.has(blockId)) return fail('Duplicate presence roots.', 'ambiguous-target');
    const owner = index.byId.get(blockId),
      part = owner && pkg.parts.get(owner.partName),
      root = part && owner && findNode(part, owner.controlId);
    if (!owner || !part || root?.kind !== 'contentControl')
      return fail('A presence root has no unique native owner.', 'ambiguous-target');
    if (collectRevisionSites({ ...part, root }).length)
      return fail(
        'Presence ownership cannot absorb pending native revisions.',
        'unsupported-operation'
      );
    if (contentControlLevelOf(root) !== 'block')
      return fail('Presence proposals require a block content control.', 'unsupported-operation');
    const members: BlockPresenceMember[] = [];
    for (const node of [
      root,
      ...contentControlsIn(root)
        .map((entry) => entry.node)
        .filter((node) => node.id !== root.id),
    ]) {
      const tag = contentControlPropertiesOf(node).tag ?? '',
        id = blockIdFromTag(tag);
      if (!id) {
        if (tag.startsWith(BLOCK_TAG_PREFIX))
          return fail('A presence subtree has a malformed block reference.');
        continue;
      }
      const child = index.byId.get(id),
        entry = payloads.entries.get(id);
      if (
        !child ||
        child.controlId !== node.id ||
        child.partName !== part.name ||
        !entry ||
        entry.decoded.status !== 'supported'
      )
        return fail(
          'A presence subtree has missing or duplicated payload ownership.',
          'ambiguous-target'
        );
      if (allMembers.has(id)) return fail('Presence roots overlap.', 'ambiguous-target');
      allMembers.add(id);
      members.push({ payload: child.payload, label: entry.label });
    }
    if (!members.some((member) => member.payload.id === blockId))
      return fail('The presence root payload is missing.');
    const nativeFingerprint = fingerprint(root);
    if (!nativeFingerprint)
      return fail('The presence subtree exceeds its native proof limits.', 'unsupported-operation');
    owners.set(blockId, {
      blockId,
      controlId: root.id,
      partName: part.name,
      members,
      nativeFingerprint: nativeFingerprint as `sha256:${string}`,
    });
  }
  return { success: true, owners };
}
