import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ExecResult } from '@docx-editor.dev/core/contracts/editor';
import type { BlockFragmentRecord, TableCellAddress } from '@docx-editor.dev/core/layout';
import { applyTreeOp, validateTreeOp, findNode, type TreeDocOp } from '@docx-editor.dev/core/store';
import { nativeTableAnchor } from './tables';

export const tableRefusal = (
  reason: string,
  code: 'locked' | 'invalidArgs' | 'unsupported' = 'invalidArgs'
): ExecResult => ({ ok: false, code, reason });

/** Read native placement, retaining nested table identities and ignoring repeated page headers. */
export function manualTableAddresses(editor: DocxEditorInstance, tableId: string) {
  const found = new Map<string, TableCellAddress & { readonly logicalGridColumn: number }>();
  const visit = (blocks: readonly BlockFragmentRecord[]) => {
    for (const block of blocks)
      if (block.kind === 'table')
        for (const row of block.rows) {
          if (row.isHeaderRepeat) continue;
          for (const cell of row.cells) {
            if (block.tableId === tableId && !found.has(cell.id))
              found.set(cell.id, {
                tableId,
                rowId: row.id,
                cellId: cell.id,
                rowIndex: row.rowIndex,
                gridColumn: cell.gridColumn,
                gridSpan: cell.gridSpan,
                logicalGridColumn: cell.logicalGridColumn ?? cell.gridColumn,
              });
            visit(cell.blocks);
          }
        }
  };
  for (const page of editor.surface?.layout().pages ?? []) {
    visit(page.fragments);
    visit(page.header?.fragments ?? []);
    visit(page.footer?.fragments ?? []);
  }
  return found;
}
export function manualTableContext(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) return null;
  const anchor = nativeTableAnchor(editor);
  if (!anchor) return null;
  const part = surface.session.partFor(surface.storyScope()) ?? surface.session.part();
  const table = findNode(part, anchor.tableId);
  if (table?.kind !== 'table') return null;
  const address = manualTableAddresses(editor, table.id).get(anchor.cellId);
  return address ? { surface, part, table, address } : null;
}
/** Preview only; actual execution repeats native mode/lock gates in the surface transaction. */
export function canManualTableOps(
  editor: DocxEditorInstance,
  ops: readonly TreeDocOp[],
  structural = false
): ExecResult {
  const s = editor.surface;
  if (!s || !editor.snapshot().editable || s.editingMode() === 'view')
    return tableRefusal('The document is open for viewing.', 'locked');
  if (structural && s.editingMode() === 'suggest')
    return tableRefusal('Stop suggesting to merge or split cells.', 'unsupported');
  let part = s.session.partFor(s.storyScope()) ?? s.session.part();
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i]!;
    const refusal = validateTreeOp(part, op);
    if (refusal) return tableRefusal(refusal);
    if (i + 1 < ops.length) {
      const preview = applyTreeOp(part, op);
      if (!preview.ok) return tableRefusal(preview.reason);
      part = preview.part;
    }
  }
  return { ok: true, changed: false };
}
export function applyManualTableOps(
  editor: DocxEditorInstance,
  ops: readonly TreeDocOp[],
  structural = false
): ExecResult {
  editor.surface?.flushPendingInput();
  const allowed = canManualTableOps(editor, ops, structural);
  if (!allowed.ok) return allowed;
  return editor.surface!.applyTableCommandPlan({
    ok: true,
    ops,
    selection: { kind: structural ? 'adoptCommittedCaret' : 'preserveSelection' },
  });
}
