import { cellSelectionBetween } from '@docx-editor.dev/core/layout';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { manualTableAddresses } from '../../src/chrome/table-native';
import { columnShape } from './manual-table-columns';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';

export function selectTableCells(
  editor: DocxEditorInstance,
  rows: readonly [number, number],
  columns: readonly [number, number]
) {
  const s = editor.surface!,
    shape = columnShape(s.session.part());
  const first = shape.rows[rows[0]]!.cells[columns[0]]!,
    last = shape.rows[rows[1]]!.cells[columns[1]]!;
  const point = { paragraphId: first.paragraphs[0]!.id, offset: 0 };
  s.setSelection({ anchor: point, head: point });
  const table = s
    .layout()
    .pages.flatMap((p) => p.fragments)
    .find((f) => f.kind === 'table')!;
  if (table.kind !== 'table') throw new Error('Missing table');
  const addresses = manualTableAddresses(editor, table.tableId),
    a = addresses.get(first.id)!,
    b = addresses.get(last.id)!;
  const selected = cellSelectionBetween(s.layout(), a, b);
  if (!selected) throw new Error('Missing selection');
  s.setCellSelection(selected);
  return selected;
}
export async function savedTableFingerprint(editor: DocxEditorInstance) {
  const read = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!read.ok) throw new Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
}
