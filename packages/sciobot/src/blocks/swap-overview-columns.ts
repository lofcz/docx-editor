import {
  contentControlPropertiesOf,
  findNode,
  moveNode,
  parentNodeOf,
  validateTreeOp,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { CanonicalEdit } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
import { nativeGridColumns } from '../diff/inserted-column-native';
import { buildBlockIndex } from './block-index';
import { readPresentationColumns } from './column-geometry';

/** Reorder the complete owned cells. Their identities, margins, and nested records move together. */
export function swapOverviewColumns(
  pkg: OoxmlPackage,
  edit: Extract<CanonicalEdit, { kind: 'swap-overview-columns' }>
): OoxmlPackage {
  const index = buildBlockIndex(pkg),
    overview = index.byId.get(edit.overviewId);
  const geometry =
    overview?.payload.type === 'preparation_recap' && readPresentationColumns(pkg, overview, index);
  const part = pkg.parts.get(edit.partName);
  if (
    !part ||
    !geometry ||
    overview?.partName !== edit.partName ||
    geometry.columns.length !== 2 ||
    geometry.columns[0]!.block.id !== edit.firstId ||
    geometry.columns[1]!.block.id !== edit.secondId
  )
    refuseWrite('The overview columns changed before their reorder.', 'stale-revision');
  for (const { block } of geometry.columns) {
    const reason = validateTreeOp(part, {
      op: 'removeContentControl',
      controlId: block.controlId,
      keepContent: true,
    });
    if (reason) refuseWrite('The overview column cannot move: ' + reason, 'unsupported-operation');
    for (
      let node: OoxmlNode | null = findNode(part, block.controlId);
      node;
      node = parentNodeOf(part, node.id)
    ) {
      if (node.kind !== 'contentControl') continue;
      const properties = contentControlPropertiesOf(node);
      if (
        properties.dataBinding ||
        (properties.lock !== undefined && properties.lock !== 'unlocked')
      )
        refuseWrite('The overview columns contain protected content.', 'unsupported-operation');
    }
  }
  const grid = nativeGridColumns(part, geometry.tableId);
  let current = part;
  for (const [first, second] of [
    [geometry.columns[0]!.cellId, geometry.columns[1]!.cellId],
    [grid[0]!.id, grid[1]!.id],
  ] as const) {
    const parent = parentNodeOf(current, first);
    if (!parent) refuseWrite('The overview column parent is missing.', 'invalid-document');
    const at = parent.children.findIndex((node) => node.id === first);
    const moved = moveNode(current, second, parent.id, at);
    if (!moved.ok) refuseWrite('The overview column cannot move.', 'apply-failed');
    current = moved.part;
  }
  return withPart(pkg, current);
}
