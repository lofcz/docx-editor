import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import type {
  BlockFragmentRecord,
  ParagraphFragmentRecord,
  StyleSpanRecord,
} from '@docx-editor.dev/core/layout';
import {
  bodyStoryRoot,
  contentControlPropertiesOf,
  findNode,
  paragraphOffsetIndex,
  storyParagraphs,
  type OoxmlPackage,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import { SCIOBOT_CHECKLIST_TAG } from '../diff/lists';

export interface NativeChecklistRecord {
  readonly partName: string;
  readonly paragraphId: string;
  readonly controlId: string;
  readonly checked: boolean;
  readonly writable: boolean;
  readonly start: number;
  readonly end: number;
  level: number;
  inCell: boolean;
}
const LIMIT = 100_000;

/** Published visibility/ranges stay authoritative. Only application-owned inline markers qualify. */
export function checklistRecords(
  pkg: OoxmlPackage,
  layout: ExportSemanticLayout
): Map<string, NativeChecklistRecord> {
  const part = pkg.parts.get(pkg.mainDocumentPart),
    body = part && bodyStoryRoot(part);
  if (!part || !body) throw new Error('The checklist projection has no canonical body.');
  const bodyIds = new Set(storyParagraphs(body).map((node) => node.id));
  const records = new Map<string, NativeChecklistRecord>();
  const deleted = new Set<string>();
  if (!layout.displayMode || layout.displayMode === 'all-markup') {
    const pending = [{ node: body, hidden: false, depth: 0 }];
    let visits = 0;
    while (pending.length) {
      const { node, hidden, depth } = pending.pop()!;
      if (++visits > LIMIT || depth > 64) throw new Error('Checklist ancestry limit exceeded.');
      const removed = hidden || node.kind === 'revisionDelete' || node.kind === 'revisionMoveFrom';
      if (removed) deleted.add(node.id);
      if (node.kind !== 'textValue')
        for (const child of node.children)
          pending.push({ node: child, hidden: removed, depth: depth + 1 });
    }
  }
  const controls = layout.contentControls ?? [];
  if (controls.length > LIMIT) throw new Error('Checklist control limit exceeded.');
  for (const control of controls) {
    if (
      control.tag !== SCIOBOT_CHECKLIST_TAG ||
      deleted.has(control.id) ||
      !control.fragments.length ||
      !control.range ||
      !bodyIds.has(control.range.paragraphId)
    )
      continue;
    const node = findNode(part, control.id),
      paragraph = findNode(part, control.range.paragraphId);
    if (!node || paragraph?.kind !== 'paragraph')
      throw new Error('A checklist control has no canonical source.');
    const props = contentControlPropertiesOf(node),
      range = paragraphOffsetIndex(paragraph as OoxmlParagraphNode).spanOf(node);
    if (
      props.tag !== SCIOBOT_CHECKLIST_TAG ||
      !props.checkbox ||
      control.level !== 'inline' ||
      !range ||
      range.end - range.start !== 1 ||
      range.start !== control.range.start ||
      range.end !== control.range.end
    )
      throw new Error(`Invalid owned checklist marker (${control.id}).`);
    if (records.has(paragraph.id))
      throw new Error(`Multiple visible checklist markers (${paragraph.id}).`);
    records.set(paragraph.id, {
      partName: part.name,
      paragraphId: paragraph.id,
      controlId: node.id,
      checked: props.checkbox.checked,
      writable:
        !control.bound && !['contentLocked', 'sdtContentLocked'].includes(control.effectiveLock),
      start: range.start,
      end: range.end,
      level: 0,
      inCell: false,
    });
  }
  return records;
}

/** Remove the marker's displayed atom from detached export records, retaining every other source offset. */
export function withoutChecklistGlyphs(
  layout: ExportSemanticLayout,
  records: Map<string, NativeChecklistRecord>
): ExportSemanticLayout {
  let visits = 0;
  const span = (value: StyleSpanRecord): readonly StyleSpanRecord[] => {
    const marker = records.get(value.range.paragraphId);
    if (!marker || value.range.end <= marker.start || value.range.start >= marker.end)
      return [value];
    if (marker.start <= value.range.start && value.range.end <= marker.end) return [];
    if (
      value.projected ||
      value.equation ||
      value.text.length !== value.range.end - value.range.start
    )
      throw new Error('A checklist atom overlaps non-exact text.');
    const parts: StyleSpanRecord[] = [];
    for (const [start, end] of [
      [value.range.start, marker.start],
      [marker.end, value.range.end],
    ])
      if (start! < end!)
        parts.push({
          ...value,
          range: { ...value.range, start: start!, end: end! },
          text: value.text.slice(start! - value.range.start, end! - value.range.start),
        });
    return parts;
  };
  const blocks = (
    values: readonly BlockFragmentRecord[],
    inCell: boolean,
    depth: number
  ): readonly BlockFragmentRecord[] => {
    if (depth > 64) throw new Error('Checklist layout depth exceeded.');
    return values.map((block) => {
      if (++visits > LIMIT) throw new Error('Checklist layout limit exceeded.');
      if (block.kind === 'table')
        return {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              blocks: blocks(cell.blocks, true, depth + 1),
            })),
          })),
        };
      const record = records.get(block.paragraphId);
      if (record) {
        const level = block.indent.left / 36 - 1;
        if (!Number.isInteger(level) || level < 0 || level > 8)
          throw new Error('An owned checklist has unsupported indentation.');
        record.level = level;
        record.inCell = inCell;
      }
      return {
        ...block,
        lines: block.lines.map((line) => ({ ...line, spans: line.spans.flatMap(span) })),
      } as ParagraphFragmentRecord;
    });
  };
  return {
    ...layout,
    pages: layout.pages.map((page) => ({ ...page, fragments: blocks(page.fragments, false, 0) })),
  };
}
