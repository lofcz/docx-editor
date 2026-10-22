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
}
