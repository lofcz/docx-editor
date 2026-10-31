import { segmentGraphemes } from '../layout/grapheme.ts';

/**
 * The UTF-16 offset one user-perceived character before `offset`. Backspace removes the whole
 * character, like Word: one code unit would split an emoji's surrogate pair, which the store
 * refuses, and a sequence such as a flag or a skin tone would lose only its last part.
 */
export function graphemeStartBefore(text: string, offset: number): number {
  if (offset <= 0) return 0;
  if (offset > text.length) return offset - 1;
  for (const segment of segmentGraphemes(text)) {
    if (segment.utf16To >= offset)
      return segment.utf16From < offset ? segment.utf16From : offset - 1;
  }
  return offset - 1;
}

/** The UTF-16 offset one user-perceived character after `offset`; the mirror for Delete. */
export function graphemeEndAfter(text: string, offset: number): number {
  if (offset >= text.length) return offset + 1;
  for (const segment of segmentGraphemes(text)) {
    if (segment.utf16To > offset) return segment.utf16To;
  }
  return offset + 1;
}
