/** The resulting paragraph a source offset belongs to: how many boundaries lie at or before it. */
export function pieceIndexOf(
  offsets: readonly number[],
  position: number,
  atBoundary: 'tail' | 'head' = 'tail'
): number {
  // Binary search — a whole-document paste can carry thousands of boundaries, and every
  // segment asks. Content sitting exactly ON a boundary opens that boundary's tail, which
  // is what the equivalent sequence of single splits does; `head` is for the range-end
  // markers the single split keeps behind, so a range that closed at the caret stays closed.
  let low = 0;
  let high = offsets.length;
  while (low < high) {
    const mid = (low + high) >> 1;
    const before = atBoundary === 'tail' ? offsets[mid]! <= position : offsets[mid]! < position;
    if (before) low = mid + 1;
    else high = mid;
  }
  return low;
}
