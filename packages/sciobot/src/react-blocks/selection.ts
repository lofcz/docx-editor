import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt, type SemanticSelection } from '@docx-editor.dev/core/layout';
import {
  contentControlContentNodeOf,
  findNode,
  storyParagraphs,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';

/** How the document selection meets one registered block. */
export type BlockSelectionRelation = 'within' | 'contains' | 'crosses';

export interface BlockSelectionEntry {
  readonly id: string;
  readonly type: string;
  readonly controlId: string;
  readonly parentId: string | null;
  readonly relation: BlockSelectionRelation;
  /** Native fragment ordinals the selected range reaches, matching React root keys. */
  readonly ordinals: readonly number[];
}

export interface BlockSelection {
  readonly selection: SemanticSelection;
  readonly collapsed: boolean;
  readonly revision: number;
  /** Outer blocks precede their descendants. Document order otherwise. */
  readonly blocks: readonly BlockSelectionEntry[];
}

const WIDGET_UI = '[data-sciobot-block-ui]';

/** True when a DOM node belongs to widget furniture rather than native document content. */
export function isWidgetSelectionNode(node: Node | null | undefined): boolean {
  const element = node?.nodeType === 1 ? (node as Element) : node?.parentElement;
  return !!element?.closest(WIDGET_UI);
}

/**
 * Read the native document selection through registered block wrappers.
 * The model selection stays authoritative. Widget form selections never enter this result.
 */
export function readBlockSelection(editor: DocxEditorInstance): BlockSelection | null {
  const surface = editor.surface;
  if (!surface) return null;
  surface.flushPendingInput();
  const selection = surface.state().selection;
  if (!selection) return null;
  const order = new Map(surface.session.paragraphIds().map((id, index) => [id, index]));
  const at = (paragraphId: string) => order.get(paragraphId);
  const anchor = at(selection.anchor.paragraphId),
    head = at(selection.head.paragraphId);
  if (anchor === undefined || head === undefined) return null;
  const forward =
    anchor < head || (anchor === head && selection.anchor.offset <= selection.head.offset);
  const start = forward ? selection.anchor : selection.head,
    end = forward ? selection.head : selection.anchor;
  const startIndex = at(start.paragraphId)!,
    endIndex = at(end.paragraphId)!;
  const collapsed = startIndex === endIndex && start.offset === end.offset;
  const layout = surface.layout();
  const pageOf = (position: SemanticSelection['anchor']) =>
    caretAt(layout, position)?.pageIndex ?? null;
  const firstPage = pageOf(start),
    lastPage = pageOf(end);
  const index = buildBlockIndex(surface.session.currentPackage());
  const boundaries = new Map(
    (layout.contentControls ?? []).map((boundary) => [boundary.id, boundary])
  );
  const story = surface.session.part();
  const entries: (BlockSelectionEntry & {
    readonly first: number;
    readonly offset: number;
    readonly depth: number;
  })[] = [];
  for (const block of index.byId.values()) {
    if (block.partName !== story.name) continue;
    const boundary = boundaries.get(block.controlId),
      node = findNode(story, block.controlId);
    if (!boundary || !node) continue;
    if (boundary.level === 'inline' && boundary.range) {
      const range = boundary.range,
        first = at(range.paragraphId);
      if (first === undefined) continue;
      const before = startIndex < first || (startIndex === first && start.offset <= range.start);
      const after = endIndex > first || (endIndex === first && end.offset >= range.end);
      const touches = collapsed
        ? startIndex === first && start.offset > range.start && start.offset < range.end
        : (startIndex < first || (startIndex === first && start.offset < range.end)) &&
          (endIndex > first || (endIndex === first && end.offset > range.start));
      if (!touches) continue;
      const within =
        startIndex === first &&
        endIndex === first &&
        start.offset >= range.start &&
        end.offset <= range.end;
      entries.push(
        Object.freeze({
          id: block.id,
          type: block.payload.type,
          controlId: block.controlId,
          parentId: block.parentId,
          relation: before && after ? 'contains' : within ? 'within' : 'crosses',
          ordinals: Object.freeze(boundary.fragments.map((_, ordinal) => ordinal)),
          first,
          offset: range.start,
          depth: boundary.nestingDepth,
        })
      );
      continue;
    }
    // Canonical ancestry owns membership. Painted wrappers only supply fragment pages.
    const positions = storyParagraphs(contentControlContentNodeOf(node) ?? node)
      .map((paragraph) => order.get(paragraph.id))
      .filter((value): value is number => value !== undefined);
    if (!positions.length) continue;
    const first = Math.min(...positions),
      last = Math.max(...positions);
    // A range ending at offset 0 of a block's first paragraph does not enter that block.
    const touchesEnd = !collapsed && endIndex === first && end.offset === 0;
    if (endIndex < first || startIndex > last || touchesEnd) continue;
    const startsInside = startIndex >= first && startIndex <= last,
      endsInside = endIndex >= first && endIndex <= last;
    const relation: BlockSelectionRelation =
      startsInside && endsInside ? 'within' : startsInside || endsInside ? 'crosses' : 'contains';
    const ordinals: number[] = [];
    boundary.fragments.forEach((fragment, ordinal) => {
      const page = fragment.pageIndex;
      if (relation === 'contains' || firstPage === null || lastPage === null)
        ordinals.push(ordinal);
      else if (page >= firstPage && page <= lastPage) ordinals.push(ordinal);
    });
    entries.push(
      Object.freeze({
        id: block.id,
        type: block.payload.type,
        controlId: block.controlId,
        parentId: block.parentId,
        relation,
        ordinals: Object.freeze(ordinals),
        first,
        offset: 0,
        depth: boundary.nestingDepth,
      })
    );
  }
  entries.sort((a, b) => a.first - b.first || a.offset - b.offset || a.depth - b.depth);
  return Object.freeze({
    selection,
    collapsed,
    revision: surface.session.packageRevision(),
    blocks: Object.freeze(
      entries.map(({ first: _first, offset: _offset, depth: _depth, ...entry }) =>
        Object.freeze(entry)
      )
    ),
  });
}
