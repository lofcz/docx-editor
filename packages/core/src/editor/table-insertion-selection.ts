import {
  findNode,
  parentNodeOf,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { SemanticSelection } from '@docx-editor.dev/core/layout';
import { readMergeableTable } from '../store/store/tree-op-table-cell-shape.ts';
import type { TreeDocxSessionView } from '../binding/tree-session.ts';

/** Capture the operation hint once and always remove the temporary subscription. */
export function captureTableCaret(session: Pick<TreeDocxSessionView, 'subscribe'>) {
  let hint: { readonly paragraphId: string; readonly start: number } | null = null;
  const dispose = session.subscribe((change) => {
    if (change.caret) hint = change.caret;
  });
  return { read: () => hint, dispose };
}

/**
 * A shared session can publish a package replacement without the local operation's
 * caret hint. The operation still names its exact following paragraph. Resolve the
 * inserted table from that canonical sibling address, never from text or paint order.
 */
export function tableInsertionSelection(
  hint: { readonly paragraphId: string; readonly start: number } | null,
  part: OoxmlPart,
  followingParagraphId: string,
  committed: boolean
): SemanticSelection | null {
  if (!committed) return null;
  if (hint && findNode(part, hint.paragraphId)?.kind === 'paragraph') {
    const point = { paragraphId: hint.paragraphId, offset: hint.start };
    return { anchor: point, head: point };
  }
  if (findNode(part, followingParagraphId)?.kind !== 'paragraph') return null;
  const parent = parentNodeOf(part, followingParagraphId);
  const index = parent?.children.findIndex((node) => node.id === followingParagraphId) ?? -1;
  const table = parent?.children[index - 1];
  if (table?.kind !== 'table') return null;
  const row = table.children.find((node) => node.kind === 'tableRow');
  const cell =
    row?.kind === 'tableRow' ? row.children.find((node) => node.kind === 'tableCell') : undefined;
  const paragraph =
    cell?.kind === 'tableCell'
      ? cell.children.find((node) => node.kind === 'paragraph')
      : undefined;
  if (!paragraph) return null;
  const point = { paragraphId: paragraph.id, offset: 0 };
  return { anchor: point, head: point };
}

/** A row command carries its exact source row and insertion side into caret resolution. */
export function tableCommandSelection(
  hint: { readonly paragraphId: string; readonly start: number } | null,
  part: OoxmlPart,
  ops: readonly TreeDocOp[]
): SemanticSelection | null {
  if (hint && findNode(part, hint.paragraphId)?.kind === 'paragraph') {
    const point = { paragraphId: hint.paragraphId, offset: hint.start };
    return { anchor: point, head: point };
  }
  const inserted = ops.find((op) => op.op === 'insertTableRow');
  if (!inserted || inserted.op !== 'insertTableRow') return columnInsertionSelection(part, ops);
  const table = findNode(part, inserted.tableId);
  if (table?.kind !== 'table') return null;
  const index = table.children.findIndex((node) => node.id === inserted.rowId);
  if (index < 0) return null;
  const row = table.children[index + (inserted.where === 'above' ? -1 : 1)];
  const cell =
    row?.kind === 'tableRow' ? row.children.find((node) => node.kind === 'tableCell') : undefined;
  const paragraph =
    cell?.kind === 'tableCell'
      ? cell.children.find((node) => node.kind === 'paragraph')
      : undefined;
  if (!paragraph) return null;
  const point = { paragraphId: paragraph.id, offset: 0 };
  return { anchor: point, head: point };
}

/** Resolve a newly allocated column cell by its retained canonical grid address. */
function columnInsertionSelection(
  part: OoxmlPart,
  ops: readonly TreeDocOp[]
): SemanticSelection | null {
  const inserted = ops.find((op) => op.op === 'insertTableColumn');
  if (!inserted || inserted.op !== 'insertTableColumn') return null;
  const read = readMergeableTable(part, inserted.tableId);
  if (!read.ok) return null;
  const { table } = read;
  let column: number;
  if ('gridColumnId' in inserted) {
    const source = table.topology.gridColumns.findIndex((col) => col.id === inserted.gridColumnId);
    if (source < 0) return null;
    column = source + (inserted.where === 'left' ? -1 : 1);
  } else {
    const source = table.index.byId.get(inserted.referenceCellId);
    if (!source) return null;
    column = inserted.where === 'left' ? source.startCol - 1 : source.endCol + 1;
  }
  for (const cells of table.index.byRow.values()) {
    const cell = cells.find(
      (entry) => entry.startCol === column && entry.span === 1 && entry.vMergeKind !== 'continue'
    );
    const paragraph = cell?.cell.children.find((node) => node.kind === 'paragraph');
    if (paragraph) {
      const point = { paragraphId: paragraph.id, offset: 0 };
      return { anchor: point, head: point };
    }
  }
  return null;
}
