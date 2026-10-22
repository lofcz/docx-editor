import {
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  parentNodeOf,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import type { BlockProjection } from './blocks';
import type { MarkdownProjection } from './markdown';
import type { NativeTextPoint } from './text';

/** Native text positions or explicit edges of a selected application block. */
export type QuotePoint =
  | NativeTextPoint
  | { readonly blockId: string; readonly edge: 'start' | 'end' };
export interface QuoteSelection {
  readonly anchor: QuotePoint;
  readonly head: QuotePoint;
}
export type QuoteProjection = MarkdownProjection &
  Partial<Pick<BlockProjection, 'blocks' | 'fallbackParagraphIds'>>;
export interface QuoteSelectionContext {
  /** Current mount and revision, read after queued editor input settles. */
  readonly revision: RevisionToken;
  /** Same-revision body part for native fallback labels and inline atomic blocks. */
  readonly part?: OoxmlPart;
  /** Same-revision all-markup projection. Removed selections use the existing live-view fallback. */
  readonly live?: QuoteProjection;
}
export interface NativeQuote {
  readonly containsMdx: boolean;
  readonly startLine: number;
  readonly endLine: number;
  readonly text: string;
  /** Unresolved results have empty text and placeholder line numbers. Never create a quote from them. */
  readonly resolved: boolean;
}
export type QuoteSelectionResult =
  | {
      readonly success: true;
      readonly quote: NativeQuote;
      readonly view: 'working' | 'live';
      readonly revision: RevisionToken;
    }
  | NativeEditFailure;

interface LocatedPoint {
  readonly offset: number;
  readonly start: number;
  readonly end: number;
}
interface PointResult {
  readonly point: LocatedPoint | null;
  readonly hidden: boolean;
}
const missing: PointResult = { point: null, hidden: false };
const hidden: PointResult = { point: null, hidden: true };
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const splitsSurrogate = (text: string, offset: number) => {
  const before = text.charCodeAt(offset - 1),
    after = text.charCodeAt(offset);
  return before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff;
};
function validPoint(point: QuotePoint): boolean {
  if (!point || typeof point !== 'object' || Array.isArray(point)) return false;
  if ('blockId' in point)
    return (
      typeof point.blockId === 'string' &&
      !!point.blockId &&
      (point.edge === 'start' || point.edge === 'end') &&
      Object.keys(point).every((key) => key === 'blockId' || key === 'edge')
    );
  return (
    typeof point.paragraphId === 'string' &&
    !!point.paragraphId &&
    Number.isSafeInteger(point.offset) &&
    point.offset >= 0 &&
    Object.keys(point).every((key) => key === 'paragraphId' || key === 'offset')
  );
}

/** Native offsets invert sparse maps only across affine text intervals or explicit boundaries. */
function textOffsets(projection: QuoteProjection, point: NativeTextPoint): number[] {
  const offsets: number[] = [];
  for (const source of projection.sourceMap.sources) {
    if (
      source.sourceScope !== 'body' ||
      source.paragraphId !== point.paragraphId ||
      !source.exact ||
      point.offset < source.sourceStart ||
      point.offset > source.sourceEnd
    )
      continue;
    if (point.offset === source.sourceStart) offsets.push(source.markdownStart);
    if (point.offset === source.sourceEnd) offsets.push(source.markdownEnd);
    const boundaries = source.markdownBoundaries ?? [];
    for (let index = 0; index < boundaries.length; index++) {
      const current = boundaries[index]!,
        next = boundaries[index + 1];
      if (current.sourceOffset === point.offset) offsets.push(current.markdownOffset);
      if (
        next &&
        current.sourceOffset < point.offset &&
        point.offset < next.sourceOffset &&
        next.sourceOffset - current.sourceOffset === next.markdownOffset - current.markdownOffset
      )
        offsets.push(current.markdownOffset + point.offset - current.sourceOffset);
    }
  }
  return offsets.filter((offset) => !splitsSurrogate(projection.text, offset));
}

function blockPoint(
  projection: QuoteProjection,
  point: NativeTextPoint,
  edge: 'start' | 'end',
  part?: OoxmlPart
): LocatedPoint | null {
  if (!part) return null;
  const paragraph = findNode(part, point.paragraphId);
  if (paragraph?.kind !== 'paragraph') return null;
  const ancestors = new Set<string>();
  let parent = parentNodeOf(part, paragraph.id);
  for (let depth = 0; parent && depth < 64; depth++) {
    ancestors.add(parent.id);
    parent = parentNodeOf(part, parent.id);
  }
  const labels = projection.fallbackParagraphIds?.includes(point.paragraphId);
  let chosen: LocatedPoint | null = null;
  for (const block of projection.blocks ?? []) {
    const node = findNode(part, block.controlId);
    if (!node) continue;
    const span = paragraphOffsetIndex(paragraph).spanOf(node);
    if (
      !(span && span.start <= point.offset && point.offset <= span.end) &&
      !(labels && ancestors.has(node.id))
    )
      continue;
    if (!chosen || block.end - block.start < chosen.end - chosen.start)
      chosen = {
        start: block.start,
        end: block.end,
        offset: edge === 'start' ? block.start : block.end,
      };
  }
  return chosen;
}

function locate(
  projection: QuoteProjection,
  point: QuotePoint,
  edge: 'start' | 'end',
  part?: OoxmlPart
): PointResult {
  if ('blockId' in point) {
    const blocks = projection.blocks?.filter((block) => block.id === point.blockId) ?? [];
    if (blocks.length !== 1) return missing;
    const block = blocks[0]!;
    return {
      hidden: false,
      point: {
        start: block.start,
        end: block.end,
        offset: point.edge === 'start' ? block.start : block.end,
      },
    };
  }
  const paragraphs = projection.sourceMap.paragraphs.filter(
    (paragraph) => paragraph.sourceScope === 'body' && paragraph.paragraphId === point.paragraphId
  );
  const offsets = textOffsets(projection, point);
  const paragraph = paragraphs[0];
  if (
    paragraphs.length === 1 &&
    paragraph?.markdownStart !== undefined &&
    paragraph.markdownEnd !== undefined
  ) {
    if (offsets.length)
      return {
        hidden: false,
        point: {
          start: paragraph.markdownStart,
          end: paragraph.markdownEnd,
          offset: offsets.reduce((best, offset) =>
            edge === 'start' ? Math.max(best, offset) : Math.min(best, offset)
          ),
        },
      };
    if (paragraph.sourceStart === paragraph.sourceEnd && point.offset === paragraph.sourceStart)
      return {
        hidden: false,
        point: {
          start: paragraph.markdownStart,
          end: paragraph.markdownEnd,
          offset: paragraph.markdownStart,
        },
      };
  }
  const block = blockPoint(projection, point, edge, part);
  if (block) return { point: block, hidden: false };
  return paragraphs.length || (part && findNode(part, point.paragraphId)?.kind === 'paragraph')
    ? hidden
    : missing;
}

/**
 * Resolve complete copyable Markdown lines from the same working projection used by agent reads.
 * This pure reader never changes selection, history, display mode, or the canonical package.
 * Native serializer escaping and domain indentation can differ from Plate; quote the landed text.
 * A selected deletion uses the explicit live projection. Its result identifies that fallback view.
 */
export function resolveQuoteSelection(
  working: QuoteProjection,
  selection: QuoteSelection,
  context: QuoteSelectionContext
): QuoteSelectionResult {
  const problem = jsonDataProblem(selection, { maxDepth: 4, maxNodes: 24 });
  if (
    problem ||
    !selection ||
    Object.keys(selection).some((key) => key !== 'anchor' && key !== 'head') ||
    !validPoint(selection.anchor) ||
    !validPoint(selection.head)
  )
    return fail('Invalid native quote selection.');
  for (const projection of [working, ...(context.live ? [context.live] : [])]) {
    if (!sameRevision(projection.revision, context.revision))
      return {
        success: false,
        code: 'stale-revision',
        error: 'The quote projection belongs to another document revision.',
        expectedRevision: Object.freeze({ ...projection.revision }),
        actualRevision: Object.freeze({ ...context.revision }),
      };
  }
  if (
    working.displayMode !== 'proposed' ||
    (context.live && context.live.displayMode !== 'all-markup')
  )
    return fail(
      'Quotes need a proposed projection and an all-markup fallback.',
      'invalid-document'
    );
  for (const point of [selection.anchor, selection.head]) {
    if ('blockId' in point || !context.part) continue;
    const text = paragraphTextOf(context.part, point.paragraphId);
    if (text !== null && (point.offset > text.length || splitsSurrogate(text, point.offset)))
      return fail('A quote endpoint is outside native text or splits a surrogate pair.');
  }
  let projection = working,
    view: 'working' | 'live' = 'working';
  let anchor = locate(projection, selection.anchor, 'start', context.part),
    head = locate(projection, selection.head, 'end', context.part);
  const sameTextRange =
    'paragraphId' in selection.anchor &&
    'paragraphId' in selection.head &&
    selection.anchor.paragraphId === selection.head.paragraphId;
  const collapsedHidden =
    sameTextRange &&
    selection.anchor.offset !== selection.head.offset &&
    anchor.point &&
    head.point &&
    anchor.point.offset === head.point.offset;
  if (anchor.hidden || head.hidden || collapsedHidden) {
    if (context.live) {
      projection = context.live;
      view = 'live';
      anchor = locate(projection, selection.anchor, 'start', context.part);
      head = locate(projection, selection.head, 'end', context.part);
    } else {
      anchor = missing;
      head = missing;
    }
  }
  const base = { success: true as const, view, revision: Object.freeze({ ...context.revision }) };
  if ((!anchor.point && !head.point) || anchor.hidden || head.hidden)
    return Object.freeze({
      ...base,
      quote: Object.freeze({
        containsMdx: false,
        startLine: 1,
        endLine: 1,
        text: '',
        resolved: false,
      }),
    });
  const from = anchor.point?.offset ?? head.point!.start,
    to = head.point?.offset ?? anchor.point!.end;
  const lineAt = (offset: number) => projection.text.slice(0, offset).split('\n').length;
  const startLine = lineAt(Math.min(from, to)),
    endLine = lineAt(Math.max(from, to));
  const text = projection.text
    .split('\n')
    .slice(startLine - 1, endLine)
    .join('\n');
  return Object.freeze({
    ...base,
    quote: Object.freeze({
      containsMdx: /<[a-zA-Z_][\w.-]*/.test(text),
      startLine,
      endLine,
      text,
      resolved: true,
    }),
  });
}
