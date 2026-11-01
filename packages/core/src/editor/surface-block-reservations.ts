import type { BlockInset } from '../layout/block-reservations.ts';

/** Layout-only room a host reserves for rendered widgets. */
export interface SurfaceBlockReservations {
  /**
   * Reserve extra flow height after body paragraphs, in points, keyed by canonical paragraph
   * ID. The map replaces every earlier reservation; an empty map clears them.
   *
   * A view input, not an edit: it never changes the document, never enters undo history,
   * never reaches a collaboration peer, and saving ignores it. Each session reserves for its
   * own rendering. Values round up to a quarter point; unchanged values do not relayout.
   * Paragraphs inside table cells, headers, footers, and notes are not reserved.
   */
  setBlockReservations(reservations: ReadonlyMap<string, number>): void;
  /**
   * Indent body paragraphs and in-flow tables by extra room on each side, in points, keyed by
   * canonical node ID. The map replaces every earlier inset; an empty map clears them.
   *
   * A view input with the same guarantees as {@link setBlockReservations}: a host uses it to
   * give a rendered container padding, so its native text sits inside the container's frame.
   * Values round up to a quarter point and are capped; an inset that leaves less than one
   * inch of column is ignored. Paragraphs inside table cells, headers, footers, notes, and
   * framed paragraphs are not inset.
   */
  setBlockInsets(insets: ReadonlyMap<string, BlockInset>): void;
}
