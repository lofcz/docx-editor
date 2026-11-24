import type { LineRecord, StyleSpanRecord } from '../layout/semantic-records.ts';

/** A word boundary can share a DOM range only when the browser can paint its natural advance. */
function plainText(span: StyleSpanRecord): boolean {
  return (
    !span.projected &&
    !span.unsupported &&
    !span.noteSeparator &&
    !span.equation &&
    !span.inlineWidget &&
    !span.fieldAtom &&
    !span.link &&
    !span.revisions?.length &&
    !span.changeSites?.length &&
    !span.syntheticBold &&
    !span.lineEndWhitespace &&
    !span.wrapAdvanceBefore &&
    span.glyphOffsetPt === undefined &&
    span.text.length === span.range.end - span.range.start &&
    /^[\x20-\x7e]+$/.test(span.text)
  );
}

/** Layout retains each word for hit testing. Plain, contiguous words share one painted text node. */
export function plainLineForPaint(line: LineRecord): LineRecord {
  const first = line.spans[0];
  if (!first || line.spans.length < 2 || line.drawings?.length || line.hangingMarkBlank)
    return line;
  const { style } = first;
  if (
    style.bold ||
    style.italic ||
    style.underline ||
    style.strike ||
    style.doubleStrike ||
    style.shading ||
    style.highlight ||
    style.shaping ||
    style.horizontalScalePercent !== 100 ||
    style.characterSpacingPt !== 0 ||
    style.border ||
    style.ligatures ||
    !plainText(first)
  )
    return line;
  let previous = first;
  for (let index = 1; index < line.spans.length; index++) {
    const span = line.spans[index]!;
    if (
      span.style !== style ||
      span.props !== first.props ||
      span.fontSlot !== first.fontSlot ||
      span.range.paragraphId !== first.range.paragraphId ||
      span.range.start !== previous.range.end ||
      span.box.height !== first.box.height ||
      span.box.y !== first.box.y ||
      Math.abs(span.box.x - (previous.box.x + previous.box.width)) > 0.001 ||
      !(previous.text.endsWith(' ') || span.text.startsWith(' ')) ||
      !plainText(span)
    )
      return line;
    previous = span;
  }
  return {
    ...line,
    spans: [
      {
        ...first,
        text: line.spans.map((span) => span.text).join(''),
        range: { ...first.range, end: previous.range.end },
        box: { ...first.box, width: previous.box.x + previous.box.width - first.box.x },
      },
    ],
  };
}
