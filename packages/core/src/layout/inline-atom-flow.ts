import type { EquationSpanRecord } from './equation-layout.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import { growLineMetrics, type PendingLine } from './pending-line.ts';
import { paragraphSpanMetadata } from './paragraph-span-metadata.ts';

/** Append one placed equation using the same ascent/descent merge as other inline atoms. */
export function appendEquationSpan(
  line: PendingLine,
  piece: FieldAwarePiece,
  paragraphId: string,
  equation: EquationSpanRecord,
  x: number
): void {
  const width = equation.geometry.box.width,
    height = equation.geometry.box.height;
  growLineMetrics(line, { height, baseline: equation.geometry.baseline });
  line.spans.push({
    range: { paragraphId, start: piece.start, end: piece.end },
    text: '\uFFFC',
    props: piece.props,
    style: piece.style,
    box: { x, y: 0, width, height },
    projected: true,
    equation,
    ...(piece.link ? { link: piece.link } : {}),
    ...paragraphSpanMetadata(piece),
  });
  line.width += width;
  line.end = piece.end;
}
