import { overlapsPendingMerge } from './table-transaction-regions.ts';
import {
  patchPropertyChild,
  CT_TCPR_SEQUENCE,
  CT_TRPR_SEQUENCE,
  CT_TBLPR_SEQUENCE,
} from './tree-op-table-properties.ts';
import { runWithoutJournalCapture } from '../package/canonical-primitive-capture.ts';
import { findNode, replaceNode, type EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { fromEdit } from './tree-op-nodes.ts';
import type { TreeDocOp, TreeOpResult } from './tree-op-types.ts';
import {
  child,
  kids,
  ownedTableSites,
  tableContainers,
  tableRecords,
  trackedTableId,
  withHiddenTableParagraphs,
} from './table-transaction-state.ts';

/**
 * Compose native-generated table histories inside their original transaction. The operation still
 * runs through native validators and appliers; its existing markers are restored by exact identity.
 * An earlier or imported table record never gains ownership from matching attribution.
 */
export function composeTrackedTableOp(
  part: OoxmlPart,
  op: TreeDocOp,
  options: EditOptions | undefined,
  apply: (part: OoxmlPart) => TreeOpResult
): TreeOpResult | null {
  const ids = options?.trackedRevisionIds,
    tableId = trackedTableId(op);
  if (!ids || !tableId) return null;
  const table = findNode(part, tableId);
  if (table?.kind !== 'table') return null;
  const records = tableRecords(table);
  if (!records.length) return null;
  const owned = ownedTableSites(ids);
  if (records.some((record) => !owned.has(record.id))) return null;
  if (overlapsPendingMerge(table, op, owned)) return { ok: false, reason: 'unsupported' };
  const recordsByContainer = new Map(
    tableContainers(table).map((container) => [
      container.id,
      kids(container).filter((node) => records.includes(node as OoxmlElement)),
    ])
  );
  const selected = new Set(records.map((record) => record.id));
  const strip = (node: OoxmlNode): OoxmlNode => {
    if (node.kind === 'textValue' || (node.kind === 'table' && node.id !== tableId)) return node;
    return {
      ...node,
      children: node.children.filter((value) => !selected.has(value.id)).map(strip),
    } as OoxmlNode;
  };
  const hidden = new Set<string>();
  const paragraphs = (node: OoxmlNode) => {
    if (node.kind === 'paragraph') hidden.add(node.id);
    for (const value of kids(node)) paragraphs(value);
  };
  for (const row of kids(table))
    if (row.kind === 'tableRow') {
      if (child(child(row, 'trPr'), 'del')) paragraphs(row);
      else
        for (const cell of kids(row))
          if (cell.kind === 'tableCell' && child(child(cell, 'tcPr'), 'cellDel')) paragraphs(cell);
    }
  const clean = runWithoutJournalCapture(() =>
    replaceNode(part, tableId, strip(table), { ...options, deferValidation: true })
  );
  if (!clean.ok) return { ok: false, reason: 'tree-invariant' };
  const result = runWithoutJournalCapture(() =>
    withHiddenTableParagraphs(ids, hidden, () => apply(clean.part))
  );
  if (!result.ok) return result;
  const after = findNode(result.part, tableId);
  if (after?.kind !== 'table') return { ok: false, reason: 'tree-invariant' };
  const restore = (node: OoxmlNode): OoxmlNode => {
    if (node.kind === 'textValue' || (node.kind === 'table' && node.id !== tableId)) return node;
    const prior = recordsByContainer.get(node.id) ?? [];
    let result = { ...node, children: node.children.map(restore) } as OoxmlElement;
    const order =
      node.localName === 'tcPr'
        ? CT_TCPR_SEQUENCE
        : node.localName === 'trPr'
          ? CT_TRPR_SEQUENCE
          : node.localName === 'tblPr'
            ? CT_TBLPR_SEQUENCE
            : ['gridCol', 'tblGridChange'];
    for (const record of prior) {
      const patched = patchPropertyChild(
        result,
        record as OoxmlElement,
        order,
        (record as OoxmlElement).localName
      );
      if (!patched.ok) throw new Error('Native table history restoration refused.');
      result = patched.container;
    }
    return result;
  };
  return fromEdit(replaceNode(part, tableId, restore(after), options), result.effect);
}
