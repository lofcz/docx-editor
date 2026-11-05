// Vertical alignment metadata for table-cell layout.

import { WML_NAMESPACE_URI, type OoxmlElement } from '@docx-editor.dev/core/store';
import type { CascadedTableFormatting } from './style-cascade.ts';

/** `w:vAlign` — where a cell's content sits when the row is taller than the content. */
export type CellVerticalAlign = 'top' | 'center' | 'bottom';

function child(node: OoxmlElement | undefined, name: string): OoxmlElement | undefined {
  for (const item of node?.children ?? []) {
    if (
      item.kind !== 'textValue' &&
      item.namespaceUri === WML_NAMESPACE_URI &&
      item.localName === name
    )
      return item;
  }
  return undefined;
}

/** Read a valid direct alignment; absent or unsupported values retain the inherited value. */
export function readCellVerticalAlign(
  cellProperties: OoxmlElement | undefined,
  fallback: CellVerticalAlign = 'top'
): CellVerticalAlign {
  const value = child(cellProperties, 'vAlign')?.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'val'
  )?.value;
  return value === 'top' || value === 'center' || value === 'bottom' ? value : fallback;
}

/** Resolve whole-table, basedOn, conditional, and direct cell alignment in cascade order. */
export function cellVerticalAlign(
  style: CascadedTableFormatting,
  conditions: readonly string[],
  direct: OoxmlElement | undefined
): CellVerticalAlign {
  let alignment: CellVerticalAlign = 'top';
  for (const properties of style.tableCellPropertyNodes ?? [])
    alignment = readCellVerticalAlign(properties, alignment);
  for (const condition of conditions) {
    const layers = style.conditionalStyleLayers?.get(condition);
    if (layers) {
      for (const layer of layers)
        alignment = readCellVerticalAlign(child(layer, 'tcPr'), alignment);
    } else {
      alignment = readCellVerticalAlign(child(style.conditional.get(condition), 'tcPr'), alignment);
    }
  }
  return readCellVerticalAlign(direct, alignment);
}
