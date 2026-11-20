import {
  findNode,
  contentControlPropertiesOf,
  contentControlsIn,
  isStoryPart,
  withoutCustomXmlNode,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { blockIdFromTag, readPayloadStore } from '../blocks/payload-opc-codec';
import { activeStorage } from '../registry/active';
import { refuseWrite, wmlElement } from '../diff/fragment-seeds';
import type { CanonicalEdit, PlannedTableOperation } from '../diff/plan';

type SiblingEdit = Extract<CanonicalEdit, { kind: 'siblings' }>;
const SIBLING_PARENTS = new Set(['body', 'tc', 'sdtContent']);

function validatedSiblingGap(part: OoxmlPart, edit: SiblingEdit): readonly string[] {
  const parent = findNode(part, edit.parentId);
  if (!parent || parent.kind === 'textValue' || !SIBLING_PARENTS.has(parent.localName))
    refuseWrite('The planned sibling parent is not a native story container.', 'ambiguous-target');
  const children = (parent as OoxmlElement).children,
    before = edit.beforeId === null ? -1 : children.findIndex((node) => node.id === edit.beforeId),
    terminal = children.findIndex((node) => wmlElement(node, 'sectPr')),
    after =
      edit.afterId === null
        ? terminal < 0
          ? children.length
          : terminal
        : children.findIndex((node) => node.id === edit.afterId);
  if ((edit.beforeId !== null && before < 0) || after < 0 || before >= after)
    refuseWrite('The original native sibling gap no longer matches the plan.', 'stale-revision');
  const gap = children.slice(before + 1, after).map((node) => node.id);
  if (gap.length !== edit.removedIds.length || gap.some((id, i) => id !== edit.removedIds[i]))
    refuseWrite('The original native sibling gap no longer matches the plan.', 'stale-revision');
  return gap;
}

const hasSection = (node: OoxmlNode): boolean =>
  wmlElement(node, 'sectPr') || (node.kind !== 'textValue' && node.children.some(hasSection));

const blockId = (node: OoxmlNode): string | null =>
  node.kind === 'contentControl'
    ? blockIdFromTag(contentControlPropertiesOf(node).tag ?? '')
    : null;

/** Collect candidate payload records before a native subtree disappears. */
function subtreePayloadIds(node: OoxmlNode): string[] {
  return [node, ...contentControlsIn(node).map((entry) => entry.node)]
    .map(blockId)
    .filter((id): id is string => id !== null);
}

/** Find payloads that a planned row or column deletion could remove. */
export function tablePayloadIds(part: OoxmlPart, operation: PlannedTableOperation): string[] {
  if (operation.op !== 'deleteTableColumn' && operation.op !== 'deleteTableRow') return [];
  const table = findNode(part, operation.tableId);
  return table ? subtreePayloadIds(table) : [];
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
export function siblingRemoval(part: OoxmlPart, edit: SiblingEdit): SiblingRemoval | string {
  const blocks = new Set<string>();
  const payloads: string[] = [];
  for (const id of validatedSiblingGap(part, edit)) {
    const node = findNode(part, id);
    const owned = node && blockId(node) !== null;
    if (
      !node ||
      (node.kind !== 'paragraph' && !wmlElement(node, 'tbl') && !owned) ||
      hasSection(node)
    )
      return 'Direct sibling removal supports complete paragraphs, tables, and application blocks without section marks.';
    if (owned) blocks.add(id);
    payloads.push(...subtreePayloadIds(node));
  }
  return { blocks, payloads };
}
