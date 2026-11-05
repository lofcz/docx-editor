import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  findNode,
  insertChildren,
  moveNode,
  removeNode,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { withoutPayloadRecords } from '../ai/direct-block-removal';
import { buildBlockIndex } from './block-index';
import { columnPercentages, columnWidthsTwips, readColumnGeometry } from './column-geometry';
import {
  columnContent,
  columnScaffold,
  columnStructureRefusal,
  writeColumnWidths,
} from './column-structure';
import { withPayloadRecord } from './payload-opc-codec';

export interface ColumnCountChange {
  readonly groupId: string;
  readonly count: number;
  readonly expectedRevision: number;
}
export type ColumnCountChangeResult =
  | { readonly success: true; readonly revision: number }
  | NativeEditFailure;
const fail = (code: NativeEditFailure['code'], error: string): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Match the application count action: keep columns from the start; gather removed content into the last survivor. */
export function setColumnCount(
  editor: DocxEditorInstance,
  request: ColumnCountChange
): ColumnCountChangeResult {
  const surface = editor.surface;
  if (!surface) return fail('document-not-ready', 'The document is not attached.');
  surface.flushPendingInput();
  if (surface.editingMode() !== 'edit')
    return fail('unsupported-operation', 'Column structure changes require editing mode.');
  if (surface.session.packageRevision() !== request.expectedRevision)
    return fail('stale-revision', 'Read the current columns before changing their count.');
  if (!Number.isInteger(request.count) || request.count < 1 || request.count > 63)
    return fail('invalid-input', 'The column count must be between 1 and 63.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.part();
  const group = buildBlockIndex(pkg).byId.get(request.groupId);
  const geometry = group && readColumnGeometry(pkg, group);
  if (
    !geometry ||
    geometry.separatorIds.size ||
    group!.partName !== part.name ||
    surface.storyScope().kind !== 'body'
  )
    return fail(
      'unsupported-operation',
      'Convert this group to complete native columns before changing its count.'
    );
  const refusal = columnStructureRefusal(part, group!.controlId);
  if (refusal) return fail('unsupported-operation', refusal);
  try {
    const width = `${100 / request.count}%`;
    const total = geometry.columns.reduce((sum, column) => sum + column.width, 0);
    const widths = columnWidthsTwips(Array(request.count).fill(width), total);
    const scale = columnPercentages(
      geometry.columns.map((column) => column.block.payload.attributes.width)
    ).reduce((sum, value) => sum + value, 0);
    // Exact authored grid weights compose across concurrent count and width edits.
    const percentages = widths.map((value) => `${Number(((value * scale) / total).toFixed(10))}%`);
    if (columnWidthsTwips(percentages, total).some((value, index) => value !== widths[index]))
      return fail(
        'unsupported-operation',
        'The stored percentage scale cannot represent these column widths.'
      );
    const original = geometry.columns.length;
    if (
      request.count === original &&
      widths.every((value, index) => value === geometry.columns[index]!.width) &&
      percentages.every(
        (value, index) => value === geometry.columns[index]!.block.payload.attributes.width
      )
    )
      return { success: true, revision: request.expectedRevision };
    const addedIds = Array.from(
      { length: Math.max(0, request.count - original) },
      () => `column-${crypto.randomUUID()}`
    );
    const change = (working: OoxmlPackage) => {
      let current = working.parts.get(part.name)!;
      if (request.count > original) {
        const scaffold = columnScaffold(current, widths.slice(original), addedIds);
        const table = findNode(current, geometry.tableId)!;
        if (table.kind !== 'table') throw Error('The column table is missing.');
        const grid = table.children.find((node) => node.localName === 'tblGrid');
        if (!grid) throw Error('The column grid is invalid.');
        const gridInsert = insertChildren(
          current,
          grid.id,
          grid.children.length,
          scaffold.grid.children
        );
        if (!gridInsert.ok) throw Error('Column grid insertion failed.');
        const row = findNode(gridInsert.part, geometry.rowId)!;
        if (row.kind !== 'tableRow') throw Error('The column row is missing.');
        const cellsInsert = insertChildren(
          gridInsert.part,
          row.id,
          row.children.length,
          scaffold.cells
        );
        if (!cellsInsert.ok) throw Error('Column cell insertion failed.');
        current = cellsInsert.part;
      } else if (request.count < original) {
        const destination = columnContent(
          current,
          geometry.columns[request.count - 1]!.block.controlId
        );
        let at = destination.children.length;
        for (const column of geometry.columns.slice(request.count)) {
          for (const child of columnContent(current, column.block.controlId).children) {
            const moved = moveNode(current, child.id, destination.id, at++);
            if (!moved.ok)
              throw Error('Removed column content could not move to its surviving column.');
            current = moved.part;
          }
        }
        // Remove the empty owned cells after their content moves. Native sizing below
        // updates the surviving grid and preferred widths in the same transaction.
        for (const column of geometry.columns.slice(request.count).reverse()) {
          const table = findNode(current, geometry.tableId)!;
          if (table.kind !== 'table') throw Error('The column table is missing.');
          const grid = table.children.find((node) => node.localName === 'tblGrid');
          if (!grid) throw Error('The column grid is invalid.');
          const cellRemoved = removeNode(current, column.cellId);
          if (!cellRemoved.ok) throw Error('Column cell removal failed.');
          const gridRemoved = removeNode(cellRemoved.part, grid.children.at(-1)!.id);
          if (!gridRemoved.ok) throw Error('Column grid removal failed.');
          current = gridRemoved.part;
        }
      }
      const sized = applyTreeOp(current, {
        op: 'setTableProperties',
        tableId: geometry.tableId,
        columnWidthsTwips: widths,
      });
      if (!sized.ok) throw Error(`Column sizing refused: ${sized.reason}.`);
      let next = withPart(working, sized.part);
      for (const [index, id] of addedIds.entries()) {
        const saved = withPayloadRecord(
          next,
          { id, version: 1, type: 'column', attributes: { width: percentages[original + index]! } },
          'column'
        );
        if (!saved.ok) throw Error(saved.reason);
        next = saved.pkg;
      }
      next = writeColumnWidths(
        next,
        geometry.columns.slice(0, request.count).map((column) => column.block),
        percentages
      );
      next = withoutPayloadRecords(
        next,
        geometry.columns.slice(request.count).map((column) => column.block.id)
      );
      if (readColumnGeometry(next, group!)?.columns.length !== request.count)
        throw Error('The resulting column count is invalid.');
      return next;
    };
    const result = surface.applyAutomationOps(() => [], { kind: 'body' }, [change], false, true);
    if (result.rejected)
      return fail('apply-failed', String(result.reason ?? 'The column count did not commit.'));
    return { success: true, revision: surface.session.packageRevision() };
  } catch (error) {
    return fail(
      'apply-failed',
      error instanceof Error ? error.message : 'The column count change failed.'
    );
  }
}
