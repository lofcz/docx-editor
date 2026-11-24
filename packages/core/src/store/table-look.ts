import type { OoxmlElement } from './package/ooxml-tree.ts';
import {
  wmlAttributeValue as attributeValue,
  wmlChildNamed as childNamed,
} from './store/tree-op-table-shared.ts';

/**
 * `w:tblLook` (17.4.55): which conditional formats of the table style are live.
 *
 * The element has two encodings: the legacy `w:val` bitmask and six named attributes
 * (`w:firstRow`, `w:lastRow`, `w:firstColumn`, `w:lastColumn`, `w:noHBand`, `w:noVBand`).
 * They are never merged:
 *
 *   - With any named attribute present, the named attributes are the whole statement and
 *     `w:val` is ignored. A named attribute the element leaves out reads as ON, so
 *     `<w:tblLook w:firstRow="1"/>` also turns on the last row, both columns, and both
 *     "no band" flags.
 *   - With only `w:val`, the bitmask decides.
 *   - With neither, including no element at all, the look is `04A0`: header row, first
 *     column and row banding.
 *
 * Measured against reference renders in compatibility modes 12, 14 and 15 with a style
 * that shades every condition differently.
 */
export interface TableLook {
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstColumn: boolean;
  readonly lastColumn: boolean;
  readonly rowBanding: boolean;
  readonly columnBanding: boolean;
}

/** `04A0`: the look of a table with no `w:tblLook`, or one that states nothing. */
const DEFAULT_TABLE_LOOK: TableLook = Object.freeze({
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  rowBanding: true,
  columnBanding: false,
});

const NAMED_LOOK_ATTRIBUTES = [
  'firstRow',
  'lastRow',
  'firstColumn',
  'lastColumn',
  'noHBand',
  'noVBand',
] as const;

function onOff(node: OoxmlElement, name: string): boolean | undefined {
  const raw = attributeValue(node, name);
  if (raw === undefined) return undefined;
  return raw !== '0' && raw !== 'false' && raw !== 'off';
}

/**
 * `w:tblLook` is read from the TABLE's own `w:tblPr` only, never cascaded from the style it
 * names. The schema admits `w:tblLook` inside a table style's `w:tblPr`, but the look
 * (ECMA-376 Part 1 §17.4.55) selects which conditional formats this table uses. It is a
 * property of this table's use of the style, not of the style itself.
 */
export function readTableLook(tblPr: OoxmlElement | undefined): TableLook {
  return tableLookOf(tblPr && childNamed(tblPr, 'tblLook'));
}

/** {@link readTableLook} for the `w:tblLook` element itself. */
export function tableLookOf(look: OoxmlElement | null | undefined): TableLook {
  if (!look) return DEFAULT_TABLE_LOOK;
  if (NAMED_LOOK_ATTRIBUTES.some((name) => attributeValue(look, name) !== undefined)) {
    // A named attribute the element leaves out reads as on.
    const named = (name: (typeof NAMED_LOOK_ATTRIBUTES)[number]): boolean =>
      onOff(look, name) ?? true;
    return {
      firstRow: named('firstRow'),
      lastRow: named('lastRow'),
      firstColumn: named('firstColumn'),
      lastColumn: named('lastColumn'),
      rowBanding: !named('noHBand'),
      columnBanding: !named('noVBand'),
    };
  }
  const rawVal = attributeValue(look, 'val');
  if (rawVal === undefined || !/^[0-9A-Fa-f]{1,4}$/.test(rawVal)) return DEFAULT_TABLE_LOOK;
  // The legacy bitmask: 0x0020 firstRow, 0x0040 lastRow, 0x0080 firstColumn,
  // 0x0100 lastColumn, 0x0200 NO row banding, 0x0400 NO column banding.
  const mask = Number.parseInt(rawVal, 16);
  return {
    firstRow: (mask & 0x0020) !== 0,
    lastRow: (mask & 0x0040) !== 0,
    firstColumn: (mask & 0x0080) !== 0,
    lastColumn: (mask & 0x0100) !== 0,
    rowBanding: (mask & 0x0200) === 0,
    columnBanding: (mask & 0x0400) === 0,
  };
}

export function tableLookAttributes(look: TableLook): Record<string, string> {
  const flags = {
    firstRow: look.firstRow,
    lastRow: look.lastRow,
    firstColumn: look.firstColumn,
    lastColumn: look.lastColumn,
    noHBand: !look.rowBanding,
    noVBand: !look.columnBanding,
  };
  let mask = 0;
  const attrs: Record<string, string> = {};
  Object.entries(flags).forEach(([name, on], index) => {
    if (on) mask |= 0x20 << index;
    attrs[name] = on ? '1' : '0';
  });
  attrs.val = mask.toString(16).toUpperCase().padStart(4, '0');
  return attrs;
}
