import type { OoxmlElement } from '@docx-editor.dev/core/store';
import type { CellVerticalAlign } from './table-cell-vertical-align.ts';
import type { CellMarginsPt } from './table-cell-margins.ts';
import type { CellBorderBox, TableBorderBox } from './table-borders.ts';
import type { PreferredWidth } from './table-widths.ts';
import type { TableCellStyleFormatting } from './style-cascade.ts';

/**
 * One cell in the resolved table structure.
 *
 * `gridSpan` is clamped at READ time and layout never re-derives it — the value comes from a file
 * and would otherwise be a loop bound an attacker controls.
 */
export interface SemanticTableCell {
  readonly revisionShading?: 'inserted' | 'deleted' | 'merged' | 'split';
  readonly revisionShadingAuthor?: string;
  readonly id: string;
  /** Derived content-edge geometry for a verified legacy percentage-width parent table. */
  readonly legacyContentAlignment?: true;
  /** Resolved simple side rules centered on a legacy absolute-width table grid. */
  readonly centeredSideRules?: true;
  /** Simple side rules painted centered on the grid line, without moving the content edge. */
  readonly centeredSidePaint?: true;
  /** Clamped to [1, MAX_TABLE_COLUMNS] at read time; layout never re-derives it. */
  readonly gridSpan: number;
  /** Physical grid column after width/style resolution and the bidiVisual projection. */
  readonly gridColumn: number;
  /** Logical column after omitted revision columns and unused boundaries are compacted. */
  readonly resolvedGridColumn?: number;
  /** Logical span in that compacted grid; physical gridSpan and geometry remain unchanged. */
  readonly resolvedGridSpan?: number;
  /** Total columns in the logical resolved grid, including preserved row gaps. */
  readonly resolvedGridColumnCount?: number;

  /** Stored grid index when bidiVisual maps this cell into a physical RTL grid. */
  readonly logicalGridColumn?: number;
  /** Canonical `w:gridCol` node id for this cell's start column, when the grid is authored. */
  readonly gridColumnId?: string;
  /** A vMerge cell that is not the restart continues the cell above: box, no content. */
  readonly vMergeContinue: boolean;
  /** This cell starts a vertical merge that the next row continues. */
  readonly headsVerticalMerge?: true;
  /** `w:hideMark` excludes the end-of-cell glyph from row sizing. */
  readonly hideEndMark?: boolean;
  /** `w:noWrap`: autofit never narrows the cell's column below its preferred width. */
  readonly noWrap?: true;
  /** `w:vAlign` — defaults to top when omitted/unrecognised. */
  readonly vAlign: CellVerticalAlign;
  /** `w:textDirection`; unsupported values keep horizontal layout. */
  readonly textDirection: 'horizontal' | 'btLr';
  /** Resolved per-side margins (tcMar over tblCellMar over the table style over Word's default). */
  readonly margins: CellMarginsPt;
  /** Three-state authored `tcBorders` (omitted / none / edge). */
  readonly borders: CellBorderBox;
  /** The row's `w:tblPrEx` table borders, which replace the table's own for this cell. */
  readonly rowTableBorders?: TableBorderBox;
  /** Resolved incident edges used for content clearance, preserving authored border provenance. */
  readonly contentBorders?: CellBorderBox;
  /** The resolved cell (including a vertical merge) ends at the authored table bottom. */
  readonly contentBottomIsOuter?: boolean;
  /** Clearance this AUTHORED ROW reserves above its content for its own top rule, in points. */
  readonly topBandClearancePt?: number;
  /** Row-local clearance in points for authored minima, before vMerge combines content boxes. */
  readonly minimumContentInsets?: { readonly top: number; readonly bottom: number };
  /** Validated 6-hex shading fill, absent for none/auto. */
  readonly shading?: string;
  /**
   * `w:tcW` — the width this cell asked for, as authored.
   *
   * Published for consumers that need the cell's own statement (a column-resize handle has
   * to write back to it). Column geometry is NOT derived from this field: the resolver works
   * from a flat claim list built in the same pass, because resolving a column means looking
   * at every cell that covers it across every row, not at one cell at a time. Read
   * `columnWidthsPt` for what the table actually laid out.
   */
  readonly preferredWidth: PreferredWidth;
  /**
   * What the table style says about this cell's paragraphs and runs (17.7.6.6) — a header
   * row's bold and centring live here, not in the cell's own properties.
   */
  readonly styleFormatting: TableCellStyleFormatting;
  /** Block children in reading order, with content-control wrappers flattened. */
  readonly blocks: readonly OoxmlElement[];
}
