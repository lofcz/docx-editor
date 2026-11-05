import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { RevisionToken, NativeEditFailure } from '../contracts/document';
import { buildBlockIndex } from '../blocks/block-index';
import {
  columnPercentages,
  columnWidthsTwips,
  readColumnGeometry,
} from '../blocks/column-geometry';
import { planDomainProperties } from './domain-properties';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';
import { planTableFormat } from './table-format';

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

/** Plan one explicit width gesture. Native grid and payload percentages share the transaction. */
export function planColumnWidths(
  source: {
    readonly pkg: OoxmlPackage;
    readonly revision: RevisionToken;
    readonly groupId: string;
  },
  widths: readonly number[]
): EditPlanResult {
  const index = buildBlockIndex(source.pkg),
    group = index.byId.get(source.groupId);
  const geometry = group && readColumnGeometry(source.pkg, group, index);
  if (!geometry) return fail('The column group needs native table geometry before resizing.');
  if (
    !Array.isArray(widths) ||
    widths.length !== geometry.columns.length ||
    widths.some((width) => !Number.isInteger(width) || width < 240 || width > 31680)
  )
    return fail('The requested column widths do not fit the native grid.');
  const total = geometry.columns.reduce((sum, column) => sum + column.width, 0);
  if (widths.reduce((sum, width) => sum + width, 0) !== total)
    return fail('Column resizing must preserve the complete group width.');
  const changed = widths.some((width, at) => width !== geometry.columns[at]!.width);
  if (!changed) return createEditPlan(source.revision, []);
  let percentages: readonly string[];
  try {
    const values = geometry.columns.map((column) => column.block.payload.attributes.width);
    const weights = columnPercentages(values),
      sum = weights.reduce((total, weight) => total + weight, 0);
    if (columnWidthsTwips(values, total).some((width, at) => width !== geometry.columns[at]!.width))
      return fail('The stored column widths differ from the native grid.');
    // Preserve the existing percentage scale and each untouched field. Normalizing every
    // field would overwrite a concurrent resize of a different boundary.
    percentages = weights.map(
      (weight, at) =>
        `${Number((weight + ((widths[at]! - geometry.columns[at]!.width) * sum) / total).toFixed(10))}%`
    );
    if (columnWidthsTwips(percentages, total).some((width, at) => width !== widths[at]))
      return fail('The requested percentages cannot preserve the exact native grid.');
  } catch (error) {
    return fail(error instanceof Error ? error.message : 'Invalid column percentages.');
  }
  const formatted = planTableFormat(
    {
      revision: source.revision,
      part: source.pkg.parts.get(group!.partName)!,
      tableId: geometry.tableId,
    },
    { op: 'setTableProperties', columnWidthsTwips: [...widths] }
  );
  if (!formatted.success) return formatted;
  const operations: CanonicalEdit[] = [...formatted.plan.operations];
  for (const [at, column] of geometry.columns.entries()) {
    if (widths[at] === column.width) continue;
    const width = percentages[at]!;
    const metadata = planDomainProperties(
      { ...source, blockId: column.block.id },
      { set: { width } },
      'field-edit'
    );
    if (!metadata.success) return metadata;
    operations.push(...metadata.plan.operations);
  }
  return createEditPlan(source.revision, operations);
}
