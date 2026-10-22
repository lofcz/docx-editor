import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import type { BlockPropertyPath } from './block-tags';
import type { BlockProjection } from './blocks';
import type { ProjectedTableCell } from './tables';
import type { ChecklistProjection } from './lists';
import { markdownRangeAt, type MarkdownProjection, type NativeMarkdownPoint } from './markdown';

export interface MarkdownPatchSpan {
  readonly start: number;
  readonly end: number;
}

/** P12 supplies protected canonical content that has no writable Markdown representation. */
export interface ProtectedMarkdownSpan extends MarkdownPatchSpan {
  readonly reason: string;
  readonly nodeId?: string;
  /** Interior-only zero-width markers allow edits beside the preserved node. Insertions at it still refuse. */
  readonly edges?: 'closed' | 'interior';
  /** P12 boundaries guard body-text edits. Generated owned properties use their own executor. */
  readonly scope?: 'body-text' | 'all';
}

export interface PatchResolutionContext {
  /** Read this from the live editor after flushing queued input. */
  readonly revision: RevisionToken;
  readonly protected?: readonly ProtectedMarkdownSpan[];
  /** P03 cell identities for this same revision; Markdown offsets can use a different table syntax. */
  readonly tableCells?: readonly Pick<ProjectedTableCell, 'id' | 'paragraphIds'>[];
}

interface PatchTargetBase {
  readonly revision: RevisionToken;
  readonly markdown: MarkdownPatchSpan;
}

export type CanonicalPatchTarget = PatchTargetBase &
  (
    | {
        readonly kind: 'text';
        readonly start: NativeMarkdownPoint;
        readonly end: NativeMarkdownPoint;
        readonly paragraphIds: readonly string[];
        readonly cellId: string | null;
      }
    | {
        readonly kind: 'property';
        readonly blockId: string;
        readonly controlId: string;
        readonly property: string;
        readonly path: BlockPropertyPath;
        readonly syntax: MarkdownPatchSpan;
        readonly value: MarkdownPatchSpan | null;
      }
    | { readonly kind: 'paragraph'; readonly paragraphId: string; readonly sourceScope: string }
    | { readonly kind: 'table'; readonly tableId: string; readonly sourceScope: string }
    | {
        readonly kind: 'checklist-state';
        readonly partName: string;
        readonly paragraphId: string;
        readonly controlId: string;
        readonly checked: boolean;
        readonly level: number;
      }
    | {
        readonly kind: 'block';
        readonly blockId: string;
        readonly controlId: string;
        readonly type: string;
      }
  );

export type PatchResolutionResult =
  | { readonly success: true; readonly target: CanonicalPatchTarget }
  | NativeEditFailure;

type MappedProjection = MarkdownProjection &
  Partial<Pick<BlockProjection, 'blocks' | 'attributes'>>;
type TaskProjection = MappedProjection & Partial<Pick<ChecklistProjection, 'checklists'>>;

function overlaps(a: MarkdownPatchSpan, b: MarkdownPatchSpan): boolean {
  if (b.start === b.end) return a.start <= b.start && b.start <= a.end;
  if (a.start === a.end) return b.start < a.start && a.start < b.end;
  return a.start < b.end && b.start < a.end;
}

function splitsSurrogate(text: string, offset: number): boolean {
  const before = text.charCodeAt(offset - 1),
    after = text.charCodeAt(offset);
  return before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff;
}

/**
 * Resolve an already located matcher span. C04 owns matching and occurrence counts.
 * This function returns addresses only. It never creates or applies canonical operations.
 * Construct targets require syntax-aware plans; they cannot become plain text replacements.
 */
export function resolvePatchSpan(
  projection: TaskProjection,
  span: MarkdownPatchSpan,
  context: PatchResolutionContext
): PatchResolutionResult {
  const fail = (error: string, code: NativeEditFailure['code'] = 'unsupported-operation') =>
    ({ success: false, code, error }) as const;
  if (!sameRevision(projection.revision, context.revision))
    return {
      success: false,
      code: 'stale-revision',
      error: 'The matched Markdown belongs to an older document revision.',
      expectedRevision: Object.freeze({ ...projection.revision }),
      actualRevision: Object.freeze({ ...context.revision }),
    };
  if (
    !Number.isSafeInteger(span.start) ||
    !Number.isSafeInteger(span.end) ||
    span.start < 0 ||
    span.end < span.start ||
    span.end > projection.text.length
  )
    return fail('The matched Markdown span is outside the projection.', 'invalid-input');
  if (splitsSurrogate(projection.text, span.start) || splitsSurrogate(projection.text, span.end))
    return fail('The matched Markdown span splits a UTF-16 surrogate pair.', 'ambiguous-target');
  const states = (projection.checklists ?? []).filter(
    (item) =>
      (span.start === item.stateStart && span.end === item.stateEnd) ||
      (span.start === item.stateStart + 1 && span.end === item.stateEnd - 1)
  );
  if (states.length > 1)
    return fail('The task state has multiple canonical identities.', 'ambiguous-target');
  const state = states[0];
  for (const region of context.protected ?? []) {
    if (
      !Number.isSafeInteger(region.start) ||
      !Number.isSafeInteger(region.end) ||
      region.start < 0 ||
      region.end < region.start ||
      region.end > projection.text.length
    )
      return fail('A protected Markdown boundary is invalid.', 'invalid-document');
    const touches =
      region.start === region.end && region.edges === 'interior'
        ? span.start === span.end
          ? span.start === region.start
          : span.start < region.start && region.start < span.end
        : overlaps(span, region);
    if (touches && !(state && region.scope === 'body-text'))
      return fail(
        `The matched Markdown touches protected content${region.nodeId ? ` (${region.nodeId})` : ''}: ${region.reason}`
      );
  }

  const base: PatchTargetBase = {
    revision: Object.freeze({ ...projection.revision }),
    markdown: Object.freeze({ start: span.start, end: span.end }),
  };
  const done = (target: CanonicalPatchTarget): PatchResolutionResult => ({
    success: true,
    target: Object.freeze(target),
  });
  if (state) {
    if (!state.writable)
      return fail(`The checklist control is locked or bound (${state.controlId}).`);
    return done({
      ...base,
      kind: 'checklist-state',
      partName: state.partName,
      paragraphId: state.paragraphId,
      controlId: state.controlId,
      checked: state.checked,
      level: state.level,
    });
  }
  const attributes = projection.attributes ?? [];
  const properties = attributes.filter(
    (property) =>
      property.start <= span.start &&
      span.end <= property.end &&
      (span.start !== span.end || span.start < property.end)
  );
  if (properties.length > 1)
    return fail(
      'The matched Markdown belongs to more than one payload property.',
      'ambiguous-target'
    );
  const property = properties[0];
  if (property)
    return done({
      ...base,
      kind: 'property',
      blockId: property.blockId,
      controlId: property.controlId,
      property: property.property,
      path: Object.freeze([...property.path]),
      syntax: Object.freeze({ start: property.start, end: property.end }),
      value:
        property.valueStart === null || property.valueEnd === null
          ? null
          : Object.freeze({ start: property.valueStart, end: property.valueEnd }),
    });

  const blocks = projection.blocks ?? [];
  const exactBlocks = blocks.filter(
    (block) => block.start === span.start && block.end === span.end
  );
  if (exactBlocks.length > 1)
    return fail('The matched construct has multiple block identities.', 'ambiguous-target');
  const block = exactBlocks[0];
  if (block)
    return done({
      ...base,
      kind: 'block',
      blockId: block.id,
      controlId: block.controlId,
      type: block.type,
    });
  if (attributes.some((attribute) => overlaps(span, attribute)))
    return fail('The matched Markdown crosses a payload property boundary.');
  for (const block of blocks) {
    if (!overlaps(span, block)) continue;
    if (
      block.contentStart === block.contentEnd ||
      span.start < block.contentStart ||
      span.end > block.contentEnd
    )
      return fail(`The matched Markdown crosses a registered block boundary (${block.id}).`);
  }

  const paragraphs = projection.sourceMap.paragraphs.filter(
    (paragraph) => paragraph.markdownStart === span.start && paragraph.markdownEnd === span.end
  );
  const tables = projection.sourceMap.tables.filter(
    (table) => table.markdownStart === span.start && table.markdownEnd === span.end
  );
  if (paragraphs.length + tables.length > 1)
    return fail('The matched construct has multiple canonical identities.', 'ambiguous-target');
  const paragraph = paragraphs[0];
  if (paragraph)
    return done({
      ...base,
      kind: 'paragraph',
      paragraphId: paragraph.paragraphId,
      sourceScope: paragraph.sourceScope,
    });
  const table = tables[0];
  if (table)
    return done({ ...base, kind: 'table', tableId: table.tableId, sourceScope: table.sourceScope });

  const range = markdownRangeAt(projection, span.start, span.end);
  if (!range)
    return fail(
      'The matched Markdown has an ambiguous or unmapped native boundary.',
      'ambiguous-target'
    );
  const ordered = projection.sourceMap.paragraphs.filter(
    (paragraph) => paragraph.sourceScope === range.start.sourceScope
  );
  const startIndex = ordered.findIndex(
    (paragraph) => paragraph.paragraphId === range.start.paragraphId
  );
  const endIndex = ordered.findIndex(
    (paragraph) => paragraph.paragraphId === range.end.paragraphId
  );
  if (startIndex < 0 || endIndex < startIndex)
    return fail('Native paragraph order is missing or ambiguous.', 'ambiguous-target');
  const paragraphIds = ordered
    .slice(startIndex, endIndex + 1)
    .map((paragraph) => paragraph.paragraphId);
  const cells = (context.tableCells ?? []).filter((cell) =>
    paragraphIds.every((id) => cell.paragraphIds.includes(id))
  );
  if (cells.length > 1) return fail('Native cell identities are ambiguous.', 'ambiguous-target');
  if (
    paragraphIds.length > 1 &&
    !cells.length &&
    projection.sourceMap.tables.some((table) =>
      overlaps(span, { start: table.markdownStart, end: table.markdownEnd })
    )
  )
    return fail('A cross-cell or unmapped table edit requires a structural table plan.');
  return done({
    ...base,
    kind: 'text',
    start: Object.freeze(range.start),
    end: Object.freeze(range.end),
    paragraphIds: Object.freeze(paragraphIds),
    cellId: cells[0]?.id ?? null,
  });
}
