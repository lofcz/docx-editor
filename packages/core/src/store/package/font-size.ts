/** Word ST_HpsMeasure: integer half-points or a positive universal length. */
export function fontSizePoints(raw: string | undefined): number | null {
  if (raw === undefined) return null;
  if (/^\d{1,7}$/.test(raw)) {
    const size = Number(raw) / 2;
    return size > 0 ? size : null;
  }
  const match = /^(\d{1,7}(?:\.\d{1,32})?)(mm|cm|in|pt|pc|pi)$/.exec(raw);
  if (!match) return null;
  const ratio: Readonly<Record<string, number>> = {
    pt: 1,
    pc: 12,
    pi: 12,
    in: 72,
    cm: 72 / 2.54,
    mm: 72 / 25.4,
  };
  const size = Number(match[1]) * ratio[match[2]!]!;
  return Number.isFinite(size) && size > 0 ? size : null;
}
/** Keep ordinary Word half-points; use exact point syntax when half-points would round. */
export function fontSizeAttribute(points: number): string {
  return Number.isInteger(points * 2) ? String(points * 2) : `${points}pt`;
}
