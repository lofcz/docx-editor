import type { CachedOutline, HarfBuzzBuffer } from './harfbuzz-runtime-types.ts';
import { fixedPoint, type FixedPoint } from './shaped-run.ts';

// HarfBuzz hb_glyph_flags_t, HB_GLYPH_FLAG_UNSAFE_TO_BREAK.
const UNSAFE_TO_BREAK = 1;

/** ECMA-376 Part 1, 17.3.2.19, w:kern (CT_HpsMeasure): keep actual positioned glyph advances. */
export function admittedGlyphWidths(
  glyphs: ReturnType<HarfBuzzBuffer['getGlyphInfos']>,
  positions: ReturnType<HarfBuzzBuffer['getGlyphPositions']>,
  known: ReadonlyMap<number, CachedOutline> | undefined,
  remainingBytes: number,
  outlineBytes: (path: string) => number,
  convert: (value: number) => FixedPoint,
  ends?: readonly number[]
): number | readonly number[] | undefined {
  if (remainingBytes < 0) return undefined;
  const widths = ends ? ([] as number[]) : undefined;
  let advance = 0;
  for (let i = 0; i < glyphs.length; i++) {
    const glyph = glyphs[i]!;
    const position = positions[i]!;
    const outline = known?.get(glyph.codepoint);
    if (
      !outline ||
      !glyph.codepoint ||
      glyph.cluster !== i ||
      position.xOffset !== 0 ||
      position.yOffset !== 0 ||
      position.yAdvance !== 0
    )
      return undefined;
    remainingBytes -= outlineBytes(outline.outline.path);
    if (remainingBytes < 0) return undefined;
    // HB_GLYPH_FLAG_UNSAFE_TO_BREAK on the following cluster requires reshaping both sides.
    if (
      ends &&
      widths &&
      widths.length &&
      i === ends[widths.length - 1] &&
      glyph.flags & UNSAFE_TO_BREAK
    )
      return undefined;
    advance = fixedPoint(advance + convert(position.xAdvance));
    if (ends && widths && i + 1 === ends[widths.length]) {
      widths.push(advance);
      advance = 0;
    }
  }
  return ends && widths ? (widths.length === ends.length ? widths : undefined) : advance;
}
