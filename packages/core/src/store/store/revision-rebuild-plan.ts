import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';
import type { Resolution } from './revision-rebuild-helpers.ts';

export interface RebuildPlan {
  /** Wrapper node id → what to do with it. */
  readonly actions: ReadonlyMap<string, Resolution>;
  readonly relocate: ReadonlyMap<string, readonly OoxmlNode[]>;
  /** Nodes to drop outright: resolved paragraph marks and spent move-range markers. */
  readonly dropMarks: ReadonlySet<string>;
  /** Inserted marks accepted for good: a container left holding nothing goes with them. */
  readonly acceptedMarks: ReadonlySet<string>;
  /** Property-change wrappers, node id → restore the recorded properties into the parent. */
  readonly restoreProperties: ReadonlySet<string>;
  /**
   * Paragraphs whose mark is resolved away, so they merge with the paragraph that follows.
   *
   * Accepting `w:pPr/w:rPr/w:del` means the paragraph mark is gone, and a paragraph without
   * its mark is not a paragraph — it runs into the next one. Removing only the `w:del`
   * element would report the deletion applied while leaving two paragraphs, which is the same
   * class of error as removing a `w:trPr/w:del` and leaving the row.
   */
  readonly mergeForward: ReadonlySet<string>;
  /** Accepting: a cell's emptied last paragraph joins the one before (rejecting keeps it). */
  readonly accepting: boolean;
  /** Removed cells, rows, and tables, including containers emptied by this decision. */
  readonly removeStructures: ReadonlySet<string>;
  readonly mergeCells: ReadonlyMap<string, OoxmlElement>;
  readonly mint: () => string;
  readonly tables: ReadonlySet<string>;
  /** The nodes on a path to a target; every other subtree is kept as it is. */
  readonly reach: ReadonlySet<string>;
}
