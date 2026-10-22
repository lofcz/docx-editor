import {
  applyTreeOp,
  directParagraphProperties,
  findNode,
  parentNodeOf,
  storyParagraphs,
  withPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { TargetAstNode } from '../projection/parse-target';
import type {
  TargetTablePropertyValues,
  TargetTablePropertyCodec,
} from './target-table-property-codec';
import type { LogicalGroup } from './logical-block-types';

export type NewCellProperties = Pick<
  TargetTablePropertyValues,
  'fill' | 'vertical' | 'horizontal' | 'borders'
>;
export interface NewCellPropertyBinding {
  readonly cellIndex: number;
  readonly properties: NewCellProperties;
}
const record = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const only = (value: object, keys: readonly string[]) =>
  Object.keys(value).every((key) => keys.includes(key));
const color = (value: unknown): boolean =>
  record(value) &&
  only(value, ['kind', 'value']) &&
  value.kind === 'hex' &&
  typeof value.value === 'string' &&
  /^[0-9A-F]{6}$/.test(value.value);
const border = (value: unknown): boolean => {
  if (!record(value)) return false;
  if (value.clear === true) return only(value, ['clear']);
  const spec = value.spec;
  return (
    value.clear === false &&
    only(value, ['clear', 'spec']) &&
    record(spec) &&
    only(spec, ['style', 'size', 'color']) &&
    ['single', 'double', 'dotted', 'dashed'].includes(String(spec.style)) &&
    Number.isSafeInteger(spec.size) &&
    Number(spec.size) >= 2 &&
    Number(spec.size) <= 96 &&
    color(spec.color)
  );
};
export function isNewCellProperties(value: unknown): value is NewCellProperties {
  if (!record(value) || !only(value, ['fill', 'vertical', 'horizontal', 'borders'])) return false;
  return (
    (!Object.hasOwn(value, 'fill') || value.fill === null || color(value.fill)) &&
    (!Object.hasOwn(value, 'vertical') ||
      ['top', 'center', 'bottom'].includes(String(value.vertical))) &&
    (!Object.hasOwn(value, 'horizontal') ||
      value.horizontal === null ||
      ['left', 'center', 'right', 'both', 'start', 'end'].includes(String(value.horizontal))) &&
    (!Object.hasOwn(value, 'borders') ||
      (record(value.borders) &&
        only(value.borders, ['outside', 'top', 'bottom', 'left', 'right']) &&
        Object.values(value.borders).every(border)))
  );
}
export function isNewCellPropertyBindings(
  cells: readonly (readonly TargetAstNode[])[],
  value: unknown
): value is readonly NewCellPropertyBinding[] {
  if (!Array.isArray(value)) return false;
  const seen = new Set<number>();
  return value.every((entry) => {
    if (
      !record(entry) ||
      !only(entry, ['cellIndex', 'properties']) ||
      !Number.isSafeInteger(entry.cellIndex) ||
      Number(entry.cellIndex) < 0 ||
      Number(entry.cellIndex) >= cells.length ||
      seen.has(Number(entry.cellIndex)) ||
      !isNewCellProperties(entry.properties)
    )
      return false;
    seen.add(Number(entry.cellIndex));
    return true;
  });
}
export const cellPropertiesAt = (
  bindings: readonly NewCellPropertyBinding[] | undefined,
  index: number
): NewCellProperties | undefined =>
  bindings?.find((entry) => entry.cellIndex === index)?.properties;

/** Resolve table/row/cell inheritance before native allocation, without retaining preview IDs. */
export function targetNewCellProperties(
  codec: TargetTablePropertyCodec,
  table: LogicalGroup,
  row: LogicalGroup,
  cell: LogicalGroup,
  column: number
): NewCellProperties {
  const tableValues = codec.values(table),
    rowValues = codec.values(row),
    values = codec.values(cell);
  const result: NewCellProperties = {};
  for (const key of ['fill', 'vertical'] as const)
    if (Object.hasOwn(values, key)) Object.assign(result, { [key]: values[key] });
  let horizontal = tableValues.horizontal;
  for (const value of [tableValues.columns?.[column], rowValues.horizontal, values.horizontal])
    if (value !== undefined && (value !== null || horizontal === undefined)) horizontal = value;
  if (horizontal !== undefined) result.horizontal = horizontal;
  if (values.borders) result.borders = values.borders;
  return result;
}

/** New cell revisions own these native properties; no preview address enters a plan. */
export function applyNewCellProperties(
  input: OoxmlPackage,
  partName: string,
  cellId: string,
  properties: NewCellProperties | undefined
): OoxmlPackage {
  if (!properties || !Object.keys(properties).length) return input;
  let pkg = input;
  const part = () => pkg.parts.get(partName)!;
  const cell = findNode(part(), cellId);
  if (cell?.kind !== 'tableCell') throw new Error('The new property cell is missing.');
  let table = parentNodeOf(part(), cellId);
  while (table && table.kind !== 'table') table = parentNodeOf(part(), table.id);
  if (!table) throw new Error('The new property cell has no table.');
  const tableId = table.id,
    cellIds = [cellId];
  const apply = (op: TreeDocOp) => {
    const result = applyTreeOp(part(), op);
    if (!result.ok) throw new Error(`New cell properties refused: ${result.reason}.`);
    pkg = withPart(pkg, result.part);
  };
  if (Object.hasOwn(properties, 'fill'))
    apply({ op: 'setTableCellFill', tableId, cellIds, color: properties.fill! });
  if (properties.vertical !== undefined)
    apply({
      op: 'setTableCellVerticalAlignment',
      tableId,
      cellIds,
      alignment: properties.vertical,
    });
  for (const side of ['outside', 'top', 'bottom', 'left', 'right'] as const) {
    const value = properties.borders?.[side];
    if (value)
      apply(
        value.clear
          ? { op: 'setTableCellBorders', tableId, cellIds, scope: 'none', target: side }
          : { op: 'setTableCellBorders', tableId, cellIds, scope: side, spec: value.spec }
      );
  }
  // R41: current new-cell seeds reject nested tables. When enabling them, scope outer
  // alignment to this cell's paragraphs so explicit inner-table alignment stays owned there.
  if (properties.horizontal !== undefined)
    for (const paragraph of storyParagraphs(cell)) {
      const existing = directParagraphProperties(part(), paragraph.id).filter(
        (entry) => entry.localName !== 'jc'
      );
      apply({
        op: 'setParagraphProperties',
        paragraphId: paragraph.id,
        properties:
          properties.horizontal === null
            ? existing
            : [...existing, { localName: 'jc', attributes: { val: properties.horizontal } }],
      });
    }
  return pkg;
}
