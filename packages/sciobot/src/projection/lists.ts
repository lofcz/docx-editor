import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { createEditPlan } from '../diff/plan';
import { checklistRecords, withoutChecklistGlyphs } from './checklist-records';
import { ProjectionComposer } from './compose';
import { projectMarkdown, type MarkdownProjection } from './markdown';
import { inlineCodeCharacterStyles } from '../native/inline-code';

export interface ProjectedChecklist {
  readonly partName: string;
  readonly paragraphId: string;
  readonly controlId: string;
  readonly checked: boolean;
  readonly level: number;
  readonly writable: boolean;
  readonly native: { readonly start: number; readonly end: number };
  readonly start: number;
  readonly end: number;
  readonly stateStart: number;
  readonly stateEnd: number;
  readonly contentStart: number;
  /** A generated NBSP keeps an empty GFM task parseable. It has no native text range. */
  readonly empty: boolean;
}
export interface ChecklistProjection extends MarkdownProjection {
  readonly checklists: readonly ProjectedChecklist[];
}

/** Translate owned checkbox state to task syntax without changing native layout, XML, or model positions. */
export function projectLists(
  pkg: OoxmlPackage,
  layout: ExportSemanticLayout,
  revision: RevisionToken,
  options: { readonly tableSyntax?: 'gfm' | 'html' | 'auto' } = {}
): { readonly success: true; readonly projection: ChecklistProjection } | NativeEditFailure {
  const token = createEditPlan(revision, []);
  if (!token.success) return token;
  try {
    const records = checklistRecords(pkg, layout);
    const trimmed = records.size ? withoutChecklistGlyphs(layout, records) : layout;
    const base = projectMarkdown(trimmed, revision, {
      codeCharacterStyles: inlineCodeCharacterStyles(pkg),
      tableSyntax: [...records.values()].some((record) => record.inCell)
        ? 'html'
        : options.tableSyntax,
    });
    if (!records.size)
      return {
        success: true,
        projection: Object.freeze({ ...base, checklists: Object.freeze([]) }),
      };
    const writer = new ProjectionComposer(base),
      checklists: ProjectedChecklist[] = [];
    const paragraphs = base.sourceMap.paragraphs
      .filter((p) => p.sourceScope === 'body' && records.has(p.paragraphId))
      .sort((a, b) => a.markdownStart! - b.markdownStart!);
    let cursor = 0,
      pendingEmpty = false;
    for (const paragraph of paragraphs) {
      const record = records.get(paragraph.paragraphId)!,
        start = paragraph.markdownStart!,
        end = paragraph.markdownEnd!;
      if (
        !Number.isSafeInteger(start) ||
        !Number.isSafeInteger(end) ||
        start < cursor ||
        end < start ||
        end > base.text.length
      )
        throw new Error('Checklist paragraph bounds overlap or are invalid.');
      if (cursor < start) {
        if (pendingEmpty && !base.text.slice(cursor, start).startsWith('\n\n'))
          writer.generated('\n\n');
        writer.native(cursor, start);
      }
      if ((start === end || pendingEmpty) && writer.offset > 0 && !writer.endsWithParagraphBreak)
        writer.generated('\n\n');
      const at = writer.offset,
        prefix = `${'  '.repeat(record.level)}- [${record.checked ? 'x' : ' '}] `;
      writer.replacement(start, start, prefix + (start === end ? '\u00a0' : ''));
      let from = start,
        newline = base.text.indexOf('\n', from);
      while (newline >= from && newline < end) {
        writer.native(from, newline + 1);
        writer.generated('  '.repeat(record.level + 1));
        from = newline + 1;
        newline = base.text.indexOf('\n', from);
      }
      writer.native(from, end);
      checklists.push(
        Object.freeze({
          partName: record.partName,
          paragraphId: record.paragraphId,
          controlId: record.controlId,
          checked: record.checked,
          level: record.level,
          writable: record.writable,
          native: Object.freeze({ start: record.start, end: record.end }),
          start: at,
          end: writer.offset,
          stateStart: at + record.level * 2 + 2,
          stateEnd: at + record.level * 2 + 5,
          contentStart: at + prefix.length,
          empty: start === end,
        })
      );
      cursor = end;
      pendingEmpty = start === end;
    }
    if (pendingEmpty && cursor < base.text.length && !base.text.slice(cursor).startsWith('\n\n'))
      writer.generated('\n\n');
    writer.native(cursor, base.text.length);
    const projected = writer.finish();
    const tasksById = new Map(checklists.map((item) => [item.paragraphId, item]));
    const paragraphBounds = base.sourceMap.paragraphs.map((paragraph) => {
      const task = paragraph.sourceScope === 'body' && tasksById.get(paragraph.paragraphId);
      if (task)
        return Object.freeze({ ...paragraph, markdownStart: task.start, markdownEnd: task.end });
      const extent =
        paragraph.markdownStart === undefined || paragraph.markdownEnd === undefined
          ? null
          : writer.mapExtent(paragraph.markdownStart, paragraph.markdownEnd, false);
      return Object.freeze({ ...paragraph, ...extent });
    });
    // An empty task still has an editable native insertion point after its marker.
    const empty = checklists
      .filter((item) => item.empty)
      .map((item) =>
        Object.freeze({
          sourceScope: 'body',
          paragraphId: item.paragraphId,
          sourceStart: item.native.end,
          sourceEnd: item.native.end,
          markdownStart: item.contentStart,
          markdownEnd: item.contentStart,
          exact: true,
          markdownBoundaries: Object.freeze([
            Object.freeze({ sourceOffset: item.native.end, markdownOffset: item.contentStart }),
          ]),
        })
      );
    return {
      success: true,
      projection: Object.freeze({
        ...projected,
        sourceMap: Object.freeze({
          ...projected.sourceMap,
          paragraphs: Object.freeze(paragraphBounds),
          sources: Object.freeze(
            [...projected.sourceMap.sources, ...empty].sort(
              (a, b) => a.markdownStart - b.markdownStart
            )
          ),
        }),
        checklists: Object.freeze(checklists),
      }),
    };
  } catch (cause) {
    return {
      success: false,
      code: 'invalid-document',
      error: cause instanceof Error ? cause.message : String(cause),
    };
  }
}

/** Keep task metadata aligned when later domain wrappers change Markdown positions. */
export function remapChecklists(
  writer: ProjectionComposer,
  items: readonly ProjectedChecklist[],
  hidden: ReadonlySet<string>
): readonly ProjectedChecklist[] {
  return Object.freeze(
    items
      .filter((item) => !hidden.has(item.paragraphId))
      .map((item) => {
        const whole = writer.mapExtent(item.start, item.end),
          state = writer.mapExtent(item.stateStart, item.stateEnd),
          point = writer.mapExtent(item.contentStart, item.contentStart);
        if (!whole || !state || !point)
          throw new Error('A task loses its source extent during domain composition.');
        return Object.freeze({
          ...item,
          start: whole.markdownStart,
          end: whole.markdownEnd,
          stateStart: state.markdownStart,
          stateEnd: state.markdownEnd,
          contentStart: point.markdownStart,
        });
      })
  );
}
