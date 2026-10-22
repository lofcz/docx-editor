import {
  findNode,
  paragraphTextOf,
  type FormattingDisplayMode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import type { TableProjection } from '../projection/tables';
import { planParagraphText } from './paragraph-text';
import { createEditPlan, type CanonicalEdit, type EditPlanResult, EDIT_PLAN_LIMITS } from './plan';
import { planRunFormat, type RunFormatTarget } from './run-format';
import { checkNativeCellScope, nearestCellAncestor, type NativeCellScope } from './cell-scope';

export interface TableCellSource extends NativeCellScope {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  /** P03 metadata from the same revision. Native layout owns grid and merge interpretation. */
  readonly projection: Pick<TableProjection, 'revision' | 'cells'>;
  readonly displayMode: FormattingDisplayMode;
}
export interface CellRunFormatPatch {
  readonly start: number;
  readonly end: number;
  readonly target: RunFormatTarget;
}
export interface CellParagraphPatch {
  readonly paragraphId: string;
  /** Complete canonical paragraph model text, with manual breaks rather than paragraph marks. */
  readonly text?: string;
  /** Original native coordinates. Formatting precedes text edits so replacement text inherits it. */
  readonly formats?: readonly CellRunFormatPatch[];
}
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/**
 * Compose text and mark plans for paragraphs belonging to one native table cell.
 * Resolve nearest native owners so nested table paragraphs cannot escape into an outer cell plan.
 * No table structural or property operations are generated. Hidden merge continuations refuse.
 * Executors must revalidate revision, ownership, old values, and protection before one transaction.
 */
export function planTableCellEdits(
  source: TableCellSource,
  patches: readonly CellParagraphPatch[]
): EditPlanResult {
  const problem = jsonDataProblem(patches, EDIT_PLAN_LIMITS);
  if (problem) return fail(problem);
  if (!Array.isArray(patches)) return fail('Cell edits need paragraph patches.');
  const raw = JSON.stringify(patches);
  if (
    raw.length > EDIT_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > EDIT_PLAN_LIMITS.maxBytes
  )
    return fail('Cell patches exceed their byte limit.');
  if (!source || !['all-markup', 'proposed', 'original'].includes(source.displayMode))
    return fail('Invalid native cell source.');
  if (!sameRevision(source.revision, source.projection.revision))
    return fail('Table projection and canonical snapshot revisions differ.', 'stale-revision');
  const records = source.projection.cells.filter((cell) => cell.id === source.cellId);
  if (
    records.length !== 1 ||
    records[0]!.tableId !== source.tableId ||
    records[0]!.rowId !== source.rowId
  )
    return fail('Table cell projection identity is missing or ambiguous.', 'ambiguous-target');
  const scopeProblem = checkNativeCellScope(source.part, {
    tableId: source.tableId,
    rowId: source.rowId,
    cellId: source.cellId,
  });
  if (scopeProblem) return scopeProblem;
  if (records[0]!.verticalContinuation)
    return fail(
      'Edit the visible merged cell instead of its hidden continuation.',
      'unsupported-operation'
    );
  const visibleParagraphIds = new Set(records[0]!.paragraphIds);
  const seen = new Set<string>(),
    formats: CanonicalEdit[] = [],
    texts: CanonicalEdit[] = [];
  for (const patch of patches) {
    if (
      !patch ||
      typeof patch !== 'object' ||
      typeof patch.paragraphId !== 'string' ||
      Object.keys(patch).some((key) => !['paragraphId', 'text', 'formats'].includes(key)) ||
      (patch.text !== undefined && typeof patch.text !== 'string') ||
      (patch.formats !== undefined && !Array.isArray(patch.formats))
    )
      return fail('Invalid cell paragraph patch.');
    if (seen.has(patch.paragraphId))
      return fail('A cell paragraph has duplicate patches.', 'ambiguous-target');
    seen.add(patch.paragraphId);
    const paragraph = findNode(source.part, patch.paragraphId);
    if (
      paragraph?.kind !== 'paragraph' ||
      !visibleParagraphIds.has(paragraph.id) ||
      nearestCellAncestor(source.part, paragraph.id, 'tableCell')?.id !== source.cellId
    )
      return fail('A paragraph does not belong to the requested visible cell.', 'ambiguous-target');
    const ranges = [...(patch.formats ?? [])];
    for (const range of ranges)
      if (
        !range ||
        typeof range !== 'object' ||
        Array.isArray(range) ||
        Object.keys(range).some((key) => !['start', 'end', 'target'].includes(key))
      )
        return fail('Invalid cell format patch.');
    ranges.sort((a, b) => a.start - b.start);
    let previousEnd = -1;
    for (const range of ranges) {
      if (range.start < previousEnd)
        return fail('Cell formatting ranges overlap.', 'ambiguous-target');
      const planned = planRunFormat(
        { ...source, paragraphId: paragraph.id, start: range.start, end: range.end },
        range.target
      );
      if (!planned.success) return planned;
      previousEnd = range.end;
      formats.push(...planned.plan.operations);
    }
    if (patch.text !== undefined) {
      const text = paragraphTextOf(source.part, paragraph.id);
      if (text === null) return fail('Cell paragraph text is unavailable.', 'invalid-document');
      const planned = planParagraphText(
        { revision: source.revision, partName: source.part.name, paragraphId: paragraph.id, text },
        patch.text
      );
      if (!planned.success) return planned;
      texts.push(...planned.plan.operations);
    }
  }
  return createEditPlan(source.revision, [...formats, ...texts]);
}
