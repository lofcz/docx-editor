import type { OoxmlPart } from '../store/package/ooxml-tree.ts';
import { readEditableTableTopology } from '../store/store/tree-op-table-topology.ts';
import { wmlAttributeValue, wmlChildNamed } from '../store/store/tree-op-table-shared.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';

/** Start repeating the first row, or stop repeating every existing header row. */
export function toggleHeaderRowOp(part: OoxmlPart, tableId: string): TreeDocOp {
  const topology = readEditableTableTopology(part.root, tableId);
  const first = topology.ok ? topology.topology.rows[0]?.row : undefined;
  const properties = first && wmlChildNamed(first, 'trPr');
  const header = properties && wmlChildNamed(properties, 'tblHeader');
  const enabled =
    header && !['0', 'false', 'off'].includes(wmlAttributeValue(header, 'val') ?? '1');
  return { op: 'setTableProperties', tableId, headerRowCount: enabled ? 0 : 1 };
}
