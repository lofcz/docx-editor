import {
  findNode,
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  withoutCustomXmlNode,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { blockIdFromTag, readPayloadStore } from '../blocks/payload-opc-codec';
import { activeStorage } from '../registry/active';
import { wmlElement } from '../diff/fragment-seeds';

const blockId = (node: OoxmlNode): string | null =>
  node.kind === 'contentControl'
    ? blockIdFromTag(contentControlPropertiesOf(node).tag ?? '')
    : null;

/**
 * The payload IDs a removed application block takes with it, nested blocks included, or null
 * when the node is not an application block.
 */
function removedBlockPayloads(node: OoxmlNode): string[] | null {
  if (blockId(node) === null) return null;
  return [node, ...contentControlsIn(node).map((entry) => entry.node)]
    .map(blockId)
    .filter((id): id is string => id !== null);
}

/** Remove the payload records of removed blocks. A record another control still uses stays. */
export function withoutPayloadRecords(pkg: OoxmlPackage, ids: readonly string[]): OoxmlPackage {
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') return pkg;
  const used = new Set<string>();
  for (const part of pkg.parts.values())
    if (isStoryPart(part))
      for (const { node } of contentControlsIn(part.root)) {
        const id = blockId(node);
        if (id !== null) used.add(id);
      }
  let next = pkg;
  for (const id of ids)
    if (store.entries.has(id) && !used.has(id))
      next = withoutCustomXmlNode(next, store.part.partName, `${activeStorage().nodePrefix}${id}`);
  return next;
}

export interface SiblingRemoval {
  /** Removed application block controls, by node ID. */
  readonly blocks: ReadonlySet<string>;
  /** Payload records the removal retires. */
  readonly payloads: readonly string[];
}

/**
 * Check direct sibling removals: complete paragraphs, tables, and application blocks without
 * section marks. Returns the refusal text, or the removed blocks and their payload IDs.
 */
export function siblingRemoval(
  part: OoxmlPart,
  removedIds: readonly string[],
  hasSection: (node: OoxmlNode) => boolean
): SiblingRemoval | string {
  const blocks = new Set<string>();
  const payloads: string[] = [];
  for (const id of removedIds) {
    const node = findNode(part, id);
    const owned = node ? removedBlockPayloads(node) : null;
    if (
      !node ||
      (node.kind !== 'paragraph' && !wmlElement(node, 'tbl') && !owned) ||
      hasSection(node)
    )
      return 'Direct sibling removal supports complete paragraphs, tables, and application blocks without section marks.';
    if (owned) {
      blocks.add(id);
      payloads.push(...owned);
    }
  }
  return { blocks, payloads };
}
