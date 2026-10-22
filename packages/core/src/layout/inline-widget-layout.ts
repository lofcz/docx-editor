import {
  contentControlsIn,
  contentControlPropertiesOf,
  paragraphOffsetIndex,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { resolveInlineWidgetBox } from '../contracts/inline-widget-metrics.ts';
import type { FieldAwarePiece } from './field-pieces.ts';
import type { TextMeasurer } from './text-measurer.ts';
import { growLineMetrics, type PendingLine } from './pending-line.ts';
import { paragraphSpanMetadata } from './paragraph-span-metadata.ts';

/** A reserved inline control box in layout points. Canonical fallback text remains unchanged. */
export interface InlineWidgetSpanRecord {
  /** Canonical inline content-control identity. */
  readonly controlId: string;
  /** Reserved advance in layout points. */
  readonly width: number;
  /** Extent above the line baseline, in points. */
  readonly ascent: number;
  /** Extent below the line baseline, in points. */
  readonly descent: number;
  /** The host width exceeded the available line width. */
  readonly clamped: boolean;
}

const providerIds = new WeakMap<object, number>();
let providerSequence = 0;
/** Include both provider identity and its explicit measurement revision in layout cache identity. */
export function inlineWidgetProducer(
  producer: string | undefined,
  measurer: TextMeasurer
): string | undefined {
  const source = measurer.inlineWidgets;
  if (!source) return producer;
  let id = providerIds.get(source.measure);
  if (id === undefined) {
    id = ++providerSequence;
    providerIds.set(source.measure, id);
  }
  return JSON.stringify([producer ?? '', 'inline-widget', id, source.revision]);
}

/** Collapse only complete visible text controls. Fields, drawings, breaks, and hidden content retain native layout. */
export function inlineWidgetPieces(
  paragraph: OoxmlNode,
  pieces: readonly FieldAwarePiece[],
  available: number,
  measurer: TextMeasurer
): readonly FieldAwarePiece[] {
  if (!measurer.inlineWidgets || paragraph.kind !== 'paragraph') return pieces;
  const index = paragraphOffsetIndex(paragraph);
  const controls = contentControlsIn(paragraph)
    .map(({ node }) => ({ node, range: index.spanOf(node) }))
    .filter((entry) => entry.range && entry.range.end > entry.range.start)
    .sort((a, b) => a.range!.start - b.range!.start || b.range!.end - a.range!.end);
  let result = pieces,
    consumed = -1;
  for (const { node, range } of controls) {
    if (!range || range.start < consumed) continue;
    const members = result.filter((piece) => piece.start < range.end && piece.end > range.start);
    const attribution = JSON.stringify([members[0]?.revisions, members[0]?.changeSites]);
    if (
      !members.length ||
      members[0]!.start !== range.start ||
      members.at(-1)!.end !== range.end ||
      members.some(
        (piece, at) =>
          piece.projected ||
          piece.fieldAtom ||
          JSON.stringify([piece.revisions, piece.changeSites]) !== attribution ||
          piece.inlineDrawing ||
          piece.equation ||
          piece.unsupported ||
          piece.breakKind ||
          piece.positionalTab ||
          /[\t\r\n]/u.test(piece.text) ||
          piece.text.length !== piece.end - piece.start ||
          (at > 0 && members[at - 1]!.end !== piece.start)
      )
    )
      continue;
    const text = members.map((piece) => piece.text).join('');
    const box = resolveInlineWidgetBox(
      {
        controlId: node.id,
        tag: contentControlPropertiesOf(node).tag,
        fallbackText: text,
        availableWidth: (available * 4) / 3,
      },
      measurer.inlineWidgets.measure
    );
    if (box.kind !== 'widget') continue;
    const widget: InlineWidgetSpanRecord = {
      controlId: node.id,
      width: Math.min(box.width * 0.75, available),
      ascent: box.ascent * 0.75,
      descent: box.descent * 0.75,
      clamped: box.clamped,
    };
    const first = members[0]!;
    const replacement: FieldAwarePiece = {
      ...first,
      text,
      start: range.start,
      end: range.end,
      projected: true,
      inlineWidget: widget,
    };
    result = result.flatMap((piece) =>
      piece === first ? [replacement] : members.includes(piece) ? [] : [piece]
    );
    consumed = range.end;
  }
  return result;
}

/** Keep measured control edges intact through native bidi itemization; declined controls stay unchanged. */
export function inlineWidgetSourceBoundaries(
  pieces: readonly FieldAwarePiece[],
  existing: ReadonlySet<number> | undefined
): ReadonlySet<number> | undefined {
  const widgets = pieces.filter((piece) => piece.inlineWidget);
  if (!widgets.length) return existing;
  return new Set([...(existing ?? []), ...widgets.flatMap((piece) => [piece.start, piece.end])]);
}

/** Restore each measured control as one atom after native text has resolved its bidi levels. */
export function restoreInlineWidgets(
  prepared: readonly FieldAwarePiece[],
  shaped: readonly FieldAwarePiece[]
): readonly FieldAwarePiece[] {
  let result = shaped;
  for (const widget of prepared.filter((piece) => piece.inlineWidget)) {
    const members = result.filter(
      (piece) => piece.start >= widget.start && piece.end <= widget.end
    );
    const first = members[0];
    if (!first || first.start !== widget.start || members.at(-1)!.end !== widget.end) continue;
    const restored = { ...widget, style: { ...widget.style, shaping: first.style.shaping } };
    result = result.flatMap((piece) =>
      piece === first ? [restored] : members.includes(piece) ? [] : [piece]
    );
  }
  return result;
}

/** Place a measured control as one unbreakable span; never rewrite its native fallback. */
export function appendInlineWidget(
  line: PendingLine,
  piece: FieldAwarePiece,
  paragraphId: string,
  x: number
): void {
  const widget = piece.inlineWidget!;
  const height = widget.ascent + widget.descent;
  growLineMetrics(line, { height, baseline: widget.ascent });
  line.spans.push({
    range: { paragraphId, start: piece.start, end: piece.end },
    text: piece.text,
    props: piece.props,
    style: piece.style,
    box: { x, y: 0, width: widget.width, height },
    projected: true,
    inlineWidget: widget,
    ...(piece.fontSlot ? { fontSlot: piece.fontSlot } : {}),
    ...(piece.link ? { link: piece.link } : {}),
    ...paragraphSpanMetadata(piece),
  });
  line.width += widget.width;
  line.end = piece.end;
}
