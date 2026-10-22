import {
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { TrackedStageWriter } from './tracked-stage-writer';

export interface StagedTextCoordinates {
  readonly runOriginScope?: string;
  /** Re-read native slices after each dependent replacement; exclude only fresh deletions. */
  slices(start: number, end: number): readonly { start: number; end: number }[];
}

/** Coordinate ownership uses exact pre-stage sites, never matching revision attribution. */
export function stagedParagraphCoordinates(
  writer: TrackedStageWriter,
  partName: string,
  paragraphId: string,
  initialSites: ReadonlySet<string>,
  expectedText?: string
): StagedTextCoordinates {
  function read() {
    const part = writer.currentPackage().parts.get(partName);
    const paragraph = part && findNode(part, paragraphId);
    if (!part || paragraph?.kind !== 'paragraph')
      refuseWrite('The native coordinate paragraph disappeared.', 'apply-failed');
    const offsets = paragraphOffsetIndex(paragraph);
    const deleted: { start: number; end: number }[] = [];
    const pending: OoxmlNode[] = [paragraph];
    while (pending.length) {
      const node = pending.pop()!;
      if (node.kind === 'revisionDelete' && !initialSites.has(node.id)) {
        const span = offsets.spanOf(node);
        if (span && span.end > span.start) deleted.push(span);
      } else if (node.kind !== 'textValue')
        for (const child of node.children) if (child.kind !== 'paragraph') pending.push(child);
    }
    deleted.sort((a, b) => a.start - b.start);
    const ranges: { start: number; end: number }[] = [];
    let cursor = 0;
    for (const span of deleted) {
      if (span.start < cursor) refuseWrite('Tracked deletion ranges overlap.', 'ambiguous-target');
      if (cursor < span.start) ranges.push({ start: cursor, end: span.start });
      cursor = span.end;
    }
    if (cursor < offsets.length) ranges.push({ start: cursor, end: offsets.length });
    const raw = paragraphTextOf(part, paragraphId)!;
    return { ranges, text: ranges.map((range) => raw.slice(range.start, range.end)).join('') };
  }
  const expected = expectedText ?? read().text;
  const checked = () => {
    const actual = read();
    if (actual.text !== expected)
      refuseWrite('Tracked text differs from its dependent coordinate model.', 'apply-failed');
    return actual.ranges;
  };
  checked();
  return {
    slices(start, end) {
      if (
        !Number.isSafeInteger(start) ||
        !Number.isSafeInteger(end) ||
        start < 0 ||
        end <= start ||
        end > expected.length
      )
        refuseWrite('A dependent text range is outside the proposed paragraph.');
      const ranges = checked(),
        result: { start: number; end: number }[] = [];
      let visible = 0;
      for (const range of ranges) {
        const length = range.end - range.start,
          from = Math.max(start, visible),
          to = Math.min(end, visible + length);
        if (from < to)
          result.push({ start: range.start + from - visible, end: range.start + to - visible });
        visible += length;
      }
      return result;
    },
  };
}
