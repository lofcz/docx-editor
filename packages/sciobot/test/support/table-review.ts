import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  AUTHORABLE_RUN_PROPERTIES,
  applyTreeOp,
  authoredProperties,
  paragraphTextOf,
  propertyContainer,
  runsCovering,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { nearestCellAncestor, type NativeCellScope } from '../../src/diff/cell-scope';
import { planTableCellEdits, type CellParagraphPatch } from '../../src/diff/table-cells';
import { directParagraphTextOps } from '../../src/diff/paragraph-text';
import type { CanonicalEditPlan } from '../../src/diff/plan';
import { readEditorRevision } from '../../src/projection/revision-cache';
import { projectTables } from '../../src/projection/tables';
import { readCommittedView } from '../../src/review/committed-view';
import { readWorkingView } from '../../src/review/working-view';
import { docxFixture } from './document';

export const CELL_CONTEXT = Object.freeze({
  documentId: 'cell-review-test',
  epoch: 'mount-1',
  author: 'ScioBot',
  date: '2026-10-19T09:00:00Z',
});
export const CELL_ORIGINAL = '| A | B |\n| --- | --- |\n| One | Two |\n\nTail';
export const cellParagraph = (text: string) =>
  `<w:p><w:r><w:rPr><w:color w:val="123456"/><w:customKeep w:val="keep"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
export function cellReviewDocument(content = cellParagraph('Two')) {
  const cell = (inner: string) => `<w:tc><w:tcPr><w:shd w:fill="EEEEEE"/></w:tcPr>${inner}</w:tc>`;
  const row = (a: string, b: string) =>
    `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(a)}${cell(b)}</w:tr>`;
  return docxFixture(
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      row(cellParagraph('A'), cellParagraph('B')) +
      row(cellParagraph('One'), content) +
      '</w:tbl>' +
      cellParagraph('Tail')
  );
}
export function cellScopeFor(
  editor: DocxEditorInstance,
  word = 'Two'
): NativeCellScope & { paragraphId: string } {
  const part = editor.surface!.session.part();
  const id = editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(part, id) === word);
  if (!id) throw new Error(`Missing table paragraph ${word}`);
  const cell = nearestCellAncestor(part, id, 'tableCell')!,
    row = nearestCellAncestor(part, cell.id, 'tableRow')!,
    table = nearestCellAncestor(part, row.id, 'table')!;
  return { cellId: cell.id, rowId: row.id, tableId: table.id, paragraphId: id };
}
export const scopeOnly = ({ cellId, rowId, tableId }: NativeCellScope): NativeCellScope => ({
  cellId,
  rowId,
  tableId,
});
export async function cellPlan(
  editor: DocxEditorInstance,
  scope: NativeCellScope,
  patches: readonly CellParagraphPatch[]
) {
  const revision = readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw new Error(view.error);
  const projected = projectTables(view.view.layout, revision);
  if (!projected.success) throw new Error(projected.error);
  const result = planTableCellEdits(
    {
      revision,
      part: editor.surface!.session.part(),
      projection: projected.projection,
      ...scopeOnly(scope),
      displayMode: 'proposed',
    },
    patches
  );
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
export async function cellViews(editor: DocxEditorInstance) {
  const revision = readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!;
  const [original, working] = await Promise.all([
    readCommittedView(editor.surface!.session, revision),
    readWorkingView(editor.surface!.session, revision),
  ]);
  if (!original.success) throw new Error(original.error);
  if (!working.success) throw new Error(working.error);
  return [original.view.projection.text, working.view.projection.text];
}
export function childNamed(
  node: OoxmlNode | null | undefined,
  name: string
): OoxmlNode | undefined {
  return node && node.kind !== 'textValue'
    ? node.children.find((child) => child.kind !== 'textValue' && child.localName === name)
    : undefined;
}
/** Pure oracle. Production writes must use the guarded surface executor. */
export function directCellResult(part: OoxmlPart, plan: CanonicalEditPlan): OoxmlPart {
  let next = part;
  for (const edit of plan.operations) {
    let ops: readonly TreeDocOp[];
    if (edit.kind === 'text') ops = directParagraphTextOps(edit);
    else if (edit.kind === 'format' && edit.target.kind === 'run') {
      const target = edit.target;
      const run = runsCovering(next, target.paragraphId, target.start, target.end)[0]!;
      const bag = new Map(
        authoredProperties(
          propertyContainer(run, 'runProperties', 'rPr'),
          AUTHORABLE_RUN_PROPERTIES
        ).map((property) => [property.localName, property])
      );
      for (const change of edit.changes) {
        if (!change.after.present) bag.delete(change.name);
        else
          bag.set(change.name, {
            localName: change.name,
            attributes: change.after.value as Record<string, string>,
          });
      }
      ops = [
        {
          op: 'setRunProperties',
          paragraphId: target.paragraphId,
          start: target.start,
          end: target.end,
          properties: [...bag.values()],
        },
      ];
    } else throw new Error('Unsupported direct cell oracle edit');
    for (const op of ops) {
      const result = applyTreeOp(next, op);
      if (!result.ok) throw new Error(result.reason);
      next = result.part;
    }
  }
  return next;
}
