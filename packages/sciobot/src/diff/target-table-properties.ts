import {
  applyTreeOp,
  findNode,
  paragraphTextOf,
  readTwipsMeasure,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import type { LogicalBlock, LogicalGroup } from './logical-block-types';
import type { TargetTableContext } from './target-table';
import type { TargetTablePropertyCodec } from './target-table-property-codec';
import type { TargetSequenceAlignment } from './target-sequence';
import type { HorizontalTarget, BorderTarget } from './target-property-values';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';
import { planTableFormat, type TableFormatTarget } from './table-format';
import { planParagraphFormat } from './paragraph-format';
import { nativeGridColumns } from './inserted-column-native';

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
const sameShape = (a: readonly LogicalBlock[], b: readonly LogicalBlock[]): boolean =>
  a.length === b.length &&
  a.every((old, index) => {
    const next = b[index]!;
    if (old.kind !== next.kind) return false;
    if (old.kind === 'paragraph' && next.kind === 'paragraph')
      return old.ast.type === next.ast.type;
    return 'children' in old && 'children' in next
      ? sameShape(old.children, next.children)
      : old.key === next.key;
  });

/** Apply explicit properties to original owners after topology planning. No preview identity can escape. */
export function planTargetTableProperties(
  context: TargetTableContext,
  codec: TargetTablePropertyCodec,
  rows: TargetSequenceAlignment['pairs'],
  columns: TargetSequenceAlignment['pairs'],
  initial: OoxmlPart,
  deferred: ReadonlySet<string> = new Set()
): EditPlanResult {
  const { input, source, before, after, bindings, beforePath } = context,
    tableId = context.binding.table!.tableId,
    revision = source.index.revision,
    table = codec.values(after),
    operations: CanonicalEdit[] = [],
    physical = input.projection.cells.filter((cell) => cell.tableId === tableId),
    byId = new Map(physical.map((cell) => [cell.id, cell])),
    paragraphs = new Map<string, HorizontalTarget>(),
    rowHeights = new Map<string, number>(),
    widths = new Map<number, number>();
  let part = initial;
  const apply = (target: TableFormatTarget): NativeEditFailure | null => {
    const result = planTableFormat({ revision, part, tableId }, target);
    if (!result.success) return result;
    for (const edit of result.plan.operations) {
      if (edit.kind !== 'table') return fail('Unexpected native table property intention.');
      const changed = applyTreeOp(part, edit.operation);
      if (!changed.ok) return fail(`Native table properties refused: ${changed.reason}.`);
      part = changed.part;
      operations.push(edit);
    }
    return null;
  };
  const selection = (id: string): readonly string[] | null => {
    const head = byId.get(id);
    if (!head || head.verticalContinuation || !findNode(part, id)) return null;
    const ids = [id];
    for (let row = head.rowIndex + 1; ; row++) {
      const next = physical.find(
        (cell) => cell.rowIndex === row && cell.gridColumn === head.gridColumn
      );
      if (!next?.verticalContinuation || next.gridSpan !== head.gridSpan) break;
      if (findNode(part, next.id)) ids.push(next.id);
    }
    return ids;
  };
  const borders = (
    ids: readonly string[],
    targets: Partial<Record<'outside' | 'top' | 'bottom' | 'left' | 'right', BorderTarget>>
  ): NativeEditFailure | null => {
    for (const side of ['outside', 'top', 'bottom', 'left', 'right'] as const) {
      const target = targets[side];
      if (!target) continue;
      const error = apply(
        target.clear
          ? { op: 'setTableCellBorders', cellIds: ids, scope: 'none', target: side }
          : { op: 'setTableCellBorders', cellIds: ids, scope: side, spec: target.spec }
      );
      if (error) return error;
    }
    return null;
  };
  const hasHorizontal =
    table.horizontal !== undefined || table.columns?.some((value) => value !== undefined);
  for (const row of rows) {
    const old = before.children[row.before] as LogicalGroup,
      next = after.children[row.after] as LogicalGroup,
      rowValues = codec.values(next),
      rowId = bindings.get([...beforePath, row.before].join('/'))?.table?.rowId;
    if (!rowId) return fail('A row property has no original native owner.');
    if (rowValues.height !== undefined) rowHeights.set(rowId, rowValues.height);
    for (const column of columns) {
      const oldCell = old.children[column.before] as LogicalGroup,
        nextCell = next.children[column.after] as LogicalGroup,
        values = codec.values(nextCell),
        owner = bindings.get([...beforePath, row.before, column.before].join('/'))?.table;
      if (deferred.has(`${row.after}/${column.after}`)) continue;
      if (!owner || owner.tableId !== tableId || !owner.cellId) {
        if (Object.keys(values).length || hasHorizontal || rowValues.horizontal !== undefined)
          return fail('A property targets a generated or missing cell.');
        continue;
      }
      const ids = selection(owner.cellId),
        head = byId.get(owner.cellId)!;
      if (values.width !== undefined) {
        if (head.gridSpan !== 1 || owner.role !== 'cell')
          return fail('A merged cell width needs explicit native span constraints.');
        if (widths.has(column.after) && widths.get(column.after) !== values.width)
          return fail('Target cells request conflicting column widths.');
        widths.set(column.after, values.width);
      }
      if (Object.hasOwn(values, 'fill') || values.vertical !== undefined || values.borders) {
        if (!ids || owner.role !== 'cell')
          return fail('Style the visible merged owner rather than a generated cell.');
        if (Object.hasOwn(values, 'fill')) {
          const error = apply({ op: 'setTableCellFill', cellIds: ids, color: values.fill! });
          if (error) return error;
        }
        if (values.vertical !== undefined) {
          const error = apply({
            op: 'setTableCellVerticalAlignment',
            cellIds: ids,
            alignment: values.vertical,
          });
          if (error) return error;
        }
        if (values.borders) {
          const error = borders(ids, values.borders);
          if (error) return error;
        }
      }
      let horizontal = table.horizontal;
      for (const inherited of [
        table.columns?.[column.after],
        rowValues.horizontal,
        values.horizontal,
      ])
        if (inherited !== undefined && (inherited !== null || horizontal === undefined))
          horizontal = inherited;
      if (horizontal !== undefined) {
        if (!sameShape(oldCell.children, nextCell.children))
          return fail('New paragraph alignment needs deferred property destinations.');
        for (const paragraphId of head.paragraphIds) {
          const paragraph = findNode(part, paragraphId),
            metadata = source.index.owners.find((owner) => owner.id === paragraphId);
          if (
            paragraph?.kind !== 'paragraph' ||
            !metadata ||
            (metadata.visibility !== 'projected' && paragraphTextOf(part, paragraphId) !== '')
          )
            return fail('Cell alignment needs explicit editable paragraph ownership.');
          if (paragraphs.has(paragraphId) && paragraphs.get(paragraphId) !== horizontal)
            return fail('Merged columns request conflicting paragraph alignments.');
          paragraphs.set(paragraphId, horizontal);
        }
      }
    }
  }
  if (table.width !== undefined || widths.size) {
    const grid = nativeGridColumns(part, tableId),
      current = grid.map((column) =>
        column.kind === 'textValue'
          ? null
          : readTwipsMeasure(
              column.attributes.find(
                (attribute) =>
                  attribute.localName === 'w' && attribute.namespaceUri === column.namespaceUri
              )?.value
            )
      );
    if (!grid.length || [...widths.keys()].some((column) => column >= grid.length))
      return fail('Table sizing needs complete native grid anchors.');
    const target = [...current],
      free = target.map((_, index) => index).filter((index) => !widths.has(index));
    for (const [index, width] of widths) target[index] = width;
    if (table.width !== undefined) {
      const remainder = table.width - [...widths.values()].reduce((a, b) => a + b, 0);
      if (remainder < 0 || (!free.length && remainder !== 0))
        return fail('Explicit table and cell widths conflict.');
      if (free.length === 1) target[free[0]!] = remainder;
      else if (free.length) {
        if (free.some((index) => current[index] === null))
          return fail('Proportional table sizing needs readable original widths.');
        const total = free.reduce((sum, index) => sum + current[index]!, 0);
        if (total <= 0) return fail('Proportional table sizing needs positive original widths.');
        let accumulated = 0,
          previous = 0;
        for (const index of free) {
          accumulated += current[index]!;
          const next = Math.round((remainder * accumulated) / total);
          target[index] = next - previous;
          previous = next;
        }
      }
    }
    if (target.some((value) => value === null))
      return fail('Table sizing needs complete target grid widths.');
    const error = apply({ op: 'setTableProperties', columnWidthsTwips: target as number[] });
    if (error) return error;
  }
  for (const [rowId, heightTwips] of rowHeights) {
    const error = apply({ op: 'setTableRowHeight', rowId, heightTwips });
    if (error) return error;
  }
  if (table.borders) {
    const ids: string[] = [],
      native = findNode(part, tableId)!;
    const visit = (node: typeof native) => {
      if (node.kind === 'textValue' || (node.kind === 'table' && node.id !== tableId)) return;
      if (node.kind === 'tableCell') ids.push(node.id);
      node.children.forEach(visit);
    };
    visit(native);
    if (ids.some((id) => !byId.has(id)))
      return fail('New table border destinations need deferred properties.');
    const error = borders(ids, table.borders);
    if (error) return error;
  }
  for (const [paragraphId, alignment] of paragraphs) {
    const result = planParagraphFormat(
      { revision, part, paragraphId },
      alignment === null
        ? { properties: [], remove: ['jc'] }
        : { properties: [{ localName: 'jc', attributes: { val: alignment } }] }
    );
    if (!result.success) return result;
    operations.push(...result.plan.operations);
  }
  return createEditPlan(revision, operations);
}
