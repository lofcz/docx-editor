import { checklistTag } from '../registry/adapter-names';
import { activeStorage } from '../registry/active';
import {
  bodyStoryRoot,
  contentControlPropertiesOf,
  paragraphOffsetIndex,
  paragraphTextOf,
  storyParagraphs,
  WML_NAMESPACE_URI,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import type { MarkdownSourceSlice } from '@docx-editor.dev/docx-to-markdown';
import { buildBlockIndex } from '../blocks/block-index';
import { sameRevision, type NativeEditFailure } from '../contracts/document';
import { copyBoundedJson } from '../data/json';
import { MDX_INLINE, MDX_BLOCK } from '../projection/mdx';
import { createEditPlan, EDIT_PLAN_LIMITS } from './plan';
import { jsonDataKey } from './target-key';
import type {
  TargetSourceInput,
  TargetSourceIndex,
  TargetSourceOwner,
  TargetSourceExtent,
  TargetSourceWindow,
  TargetSourceResult,
} from './target-source-types';

export type * from './target-source-types';
export const TARGET_SOURCE_LIMITS = Object.freeze({
  ...EDIT_PLAN_LIMITS,
  maxDepth: 96,
  maxCanonicalDepth: 256,
});
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const array = (value: unknown): boolean => Array.isArray(value);
const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
const word = (node: OoxmlNode, name: string) =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const integer = (value: unknown): value is number =>
  Number.isSafeInteger(value) && Number(value) >= 0;
function freeze(value: unknown): void {
  if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return;
  for (const child of Object.values(value)) freeze(child);
  Object.freeze(value);
}
interface Entry {
  node: OoxmlNode;
  parentId: string | null;
  childIndex: number;
  paragraphId: string | null;
  tableId: string | null;
  rowId: string | null;
  cellId: string | null;
  generated: boolean;
  atomic: boolean;
  outsideStory: boolean;
}

/**
 * Resolve projected source metadata to actual native owners and contiguous sibling windows.
 * This is a read index, not target normalization, an edit plan, or permission to mutate a node.
 * P12/A01 still validate protection, old values, live revisions, and complete preview effects.
 */
export function indexTargetSource(source: TargetSourceInput): TargetSourceResult {
  if (!source || typeof source !== 'object') return fail('Missing target source.');
  const descriptors = Object.getOwnPropertyDescriptors(source);
  for (const name of ['revision', 'pkg', 'projection'])
    if (!descriptors[name] || !('value' in descriptors[name]!) || !descriptors[name]!.enumerable)
      return fail('Target source fields must be own data values.');
  const copied = copyBoundedJson(
    { revision: source.revision, projection: source.projection },
    TARGET_SOURCE_LIMITS
  );
  if (!copied.ok) return fail(copied.reason);
  const input = copied.value as unknown as Pick<TargetSourceInput, 'revision' | 'projection'>,
    p = input.projection;
  const token = createEditPlan(input.revision, []);
  if (!token.success) return token;
  if (!record(p)) return fail('Missing MDX source projection.');
  const projectedToken = createEditPlan(p.revision, []);
  if (!projectedToken.success) return projectedToken;
  if (!sameRevision(token.plan.revision, projectedToken.plan.revision))
    return {
      ...fail('The source projection is stale.', 'stale-revision'),
      expectedRevision: projectedToken.plan.revision,
      actualRevision: token.plan.revision,
    };
  if (
    typeof p.text !== 'string' ||
    typeof p.displayMode !== 'string' ||
    !['original', 'proposed', 'all-markup'].includes(p.displayMode) ||
    !record(p.sourceMap) ||
    p.sourceMap.unit !== 'utf16-code-unit' ||
    !array(p.sourceMap.paragraphs) ||
    !array(p.sourceMap.sources) ||
    !array(p.sourceMap.tables) ||
    !array(p.blocks) ||
    !array(p.fallbackParagraphIds) ||
    !array(p.opaqueMdx) ||
    !array(p.checklists)
  )
    return fail('Invalid MDX source metadata.');
  const range = (start: unknown, end: unknown, max = p.text.length) =>
    integer(start) && integer(end) && start <= end && end <= max;
  if (
    !source.pkg ||
    typeof source.pkg.mainDocumentPart !== 'string' ||
    typeof source.pkg.parts?.get !== 'function'
  )
    return fail('Missing canonical package.', 'invalid-document');
  const part = source.pkg.parts.get(source.pkg.mainDocumentPart),
    body = part && bodyStoryRoot(part);
  if (!part || !body)
    return fail('The canonical document body does not exist.', 'invalid-document');
  const blocks = new Map<string, (typeof p.blocks)[number]>(),
    blockIds = new Set<string>();
  for (const block of p.blocks) {
    if (
      !record(block) ||
      typeof block.id !== 'string' ||
      !block.id ||
      typeof block.controlId !== 'string' ||
      !block.controlId ||
      typeof block.type !== 'string' ||
      !block.type ||
      typeof block.tag !== 'string' ||
      !block.tag ||
      !['container', 'atomic'].includes(block.kind) ||
      !range(block.start, block.end) ||
      !range(block.contentStart, block.contentEnd) ||
      block.contentStart < block.start ||
      block.contentEnd > block.end
    )
      return fail('Invalid projected block owner.');
    if (blocks.has(block.controlId) || blockIds.has(block.id))
      return fail('Duplicate projected block owner.', 'ambiguous-target');
    blocks.set(block.controlId, block);
    blockIds.add(block.id);
  }
  const entries = new Map<string, Entry>();
  const stack = [
    {
      node: body,
      parentId: null,
      childIndex: 0,
      paragraphId: null,
      tableId: null,
      rowId: null,
      cellId: null,
      generated: false,
      atomic: false,
      outsideStory: false,
      depth: 0,
    } as Entry & { depth: number },
  ];
  while (stack.length) {
    const current = stack.pop()!,
      node = current.node;
    if (
      entries.has(node.id) ||
      entries.size >= TARGET_SOURCE_LIMITS.maxNodes ||
      current.depth > TARGET_SOURCE_LIMITS.maxCanonicalDepth
    )
      return fail('The canonical body is duplicated, cyclic, or excessive.', 'invalid-document');
    const label =
      node.kind === 'contentControl' &&
      contentControlPropertiesOf(node).tag === activeStorage().labelTag;
    const entry: Entry = {
      ...current,
      paragraphId: node.kind === 'paragraph' ? node.id : current.paragraphId,
      generated: current.generated || label,
      outsideStory: current.outsideStory || word(node, 'txbxContent'),
    };
    entries.set(node.id, entry);
    const descendants = children(node);
    for (let i = descendants.length - 1; i >= 0; i--)
      stack.push({
        ...entry,
        node: descendants[i]!,
        parentId: node.id,
        childIndex: i,
        depth: current.depth + 1,
        atomic: current.atomic || blocks.get(node.id)?.kind === 'atomic',
        tableId: node.kind === 'table' ? node.id : current.tableId,
        rowId: node.kind === 'table' ? null : node.kind === 'tableRow' ? node.id : current.rowId,
        cellId:
          node.kind === 'table' || node.kind === 'tableRow'
            ? null
            : node.kind === 'tableCell'
              ? node.id
              : current.cellId,
      });
  }
  const bodyParagraphs = new Set(storyParagraphs(body).map((node) => node.id));
  const extents = new Map<string, TargetSourceExtent[]>(),
    slices = new Map<string, MarkdownSourceSlice[]>();
  const paragraph = (id: string) => {
    const entry = entries.get(id);
    return entry?.node.kind === 'paragraph' && bodyParagraphs.has(id) && !entry.outsideStory
      ? entry
      : undefined;
  };
  const textLengths = new Map<string, number>();
  const length = (id: string) => {
    if (!textLengths.has(id)) textLengths.set(id, paragraphTextOf(part, id)?.length ?? -1);
    return textLengths.get(id)!;
  };
  const add = (id: string, extent: TargetSourceExtent) => {
    const items = extents.get(id) ?? [];
    items.push(extent);
    extents.set(id, items);
  };
  for (const item of p.sourceMap.paragraphs) {
    if (!record(item) || typeof item.sourceScope !== 'string')
      return fail('Invalid paragraph source extent.');
    if (item.sourceScope !== 'body') continue;
    const owner = paragraph(item.paragraphId);
    if (!owner || !range(item.sourceStart, item.sourceEnd, length(item.paragraphId)))
      return fail('A paragraph extent has no native body owner.', 'invalid-document');
    if (owner.generated || owner.atomic)
      return fail(
        'Generated or atomic fallback cannot own projected body text.',
        'invalid-document'
      );
    if (item.markdownStart === undefined && item.markdownEnd === undefined) continue;
    if (!range(item.markdownStart, item.markdownEnd))
      return fail('Invalid paragraph Markdown bounds.');
    add(item.paragraphId, {
      start: item.markdownStart!,
      end: item.markdownEnd!,
      native: { start: item.sourceStart, end: item.sourceEnd },
    });
  }
  for (const item of p.sourceMap.sources) {
    if (!record(item) || typeof item.sourceScope !== 'string')
      return fail('Invalid inline source slice.');
    if (item.sourceScope !== 'body') continue;
    const owner = paragraph(item.paragraphId);
    if (
      !owner ||
      owner.generated ||
      owner.atomic ||
      !range(item.sourceStart, item.sourceEnd, length(item.paragraphId)) ||
      !range(item.markdownStart, item.markdownEnd) ||
      typeof item.exact !== 'boolean'
    )
      return fail('An inline slice has no valid native body owner.', 'invalid-document');
    if (item.markdownBoundaries !== undefined) {
      if (!array(item.markdownBoundaries)) return fail('Invalid inline source boundaries.');
      let native = -1,
        markdown = -1;
      for (const boundary of item.markdownBoundaries) {
        if (
          !record(boundary) ||
          !integer(boundary.sourceOffset) ||
          !integer(boundary.markdownOffset) ||
          boundary.sourceOffset < native ||
          boundary.markdownOffset < markdown ||
          boundary.sourceOffset < item.sourceStart ||
          boundary.sourceOffset > item.sourceEnd ||
          boundary.markdownOffset < item.markdownStart ||
          boundary.markdownOffset > item.markdownEnd
        )
          return fail('Invalid inline source boundary ordering.');
        native = boundary.sourceOffset;
        markdown = boundary.markdownOffset;
      }
    }
    const items = slices.get(item.paragraphId) ?? [];
    items.push(item);
    slices.set(item.paragraphId, items);
  }
  for (const item of p.sourceMap.tables) {
    if (!record(item) || typeof item.sourceScope !== 'string')
      return fail('Invalid table source extent.');
    if (item.sourceScope !== 'body') continue;
    const owner = entries.get(item.tableId);
    if (
      owner?.node.kind !== 'table' ||
      owner.outsideStory ||
      owner.generated ||
      owner.atomic ||
      !range(item.markdownStart, item.markdownEnd)
    )
      return fail('A table extent has no native body owner.', 'invalid-document');
    add(item.tableId, { start: item.markdownStart, end: item.markdownEnd, native: null });
  }
  const indexed = buildBlockIndex(source.pkg);
  for (const block of blocks.values()) {
    const owner = entries.get(block.controlId),
      native = indexed.byId.get(block.id);
    if (
      owner?.node.kind !== 'contentControl' ||
      owner.outsideStory ||
      owner.generated ||
      owner.atomic ||
      (owner.paragraphId !== null && !bodyParagraphs.has(owner.paragraphId)) ||
      !native ||
      native.controlId !== block.controlId ||
      native.partName !== part.name ||
      native.payload.type !== block.type
    )
      return fail('A projected block has no unique native payload owner.', 'invalid-document');
    add(block.controlId, { start: block.start, end: block.end, native: null });
  }
  for (const id of p.fallbackParagraphIds) {
    const owner = typeof id === 'string' ? paragraph(id) : undefined;
    if (!owner || (!owner.generated && !owner.atomic))
      return fail('A fallback paragraph has no generated or atomic owner.', 'invalid-document');
  }
  const opaque = new Map<string, (typeof p.opaqueMdx)[number]>();
  for (const item of p.opaqueMdx) {
    const block =
      record(item) && typeof item.controlId === 'string' ? blocks.get(item.controlId) : undefined;
    const payload = block && indexed.byId.get(block.id)?.payload;
    if (
      !block ||
      block.kind !== 'atomic' ||
      block.id !== item.blockId ||
      block.start !== item.start ||
      block.end !== item.end ||
      typeof item.inline !== 'boolean' ||
      !payload ||
      ![MDX_INLINE, MDX_BLOCK].includes(payload.type) ||
      item.inline !== (payload.type === MDX_INLINE) ||
      !record(item.ast) ||
      item.ast.type !== (item.inline ? 'mdxJsxTextElement' : 'mdxJsxFlowElement') ||
      jsonDataKey(item.ast) !== jsonDataKey(indexed.byId.get(block.id)!.payload.mdast!)
    )
      return fail('Opaque source metadata differs from its native payload.', 'invalid-document');
    if (opaque.has(item.controlId))
      return fail('Duplicate opaque source ownership.', 'ambiguous-target');
    opaque.set(item.controlId, item);
  }
  const tasks = new Map<string, (typeof p.checklists)[number]>();
  for (const task of p.checklists) {
    const owner =
        record(task) && typeof task.paragraphId === 'string'
          ? paragraph(task.paragraphId)
          : undefined,
      control = record(task) ? entries.get(task.controlId) : undefined;
    if (
      !owner ||
      owner.generated ||
      owner.atomic ||
      control?.node.kind !== 'contentControl' ||
      control.paragraphId !== task.paragraphId ||
      task.partName !== part.name ||
      contentControlPropertiesOf(control.node).tag !== checklistTag() ||
      !contentControlPropertiesOf(control.node).checkbox ||
      typeof task.checked !== 'boolean' ||
      typeof task.writable !== 'boolean' ||
      typeof task.empty !== 'boolean' ||
      !integer(task.level) ||
      task.level > 8 ||
      !range(task.start, task.end) ||
      !range(task.stateStart, task.stateEnd) ||
      task.stateStart < task.start ||
      task.stateEnd > task.end ||
      !integer(task.contentStart) ||
      task.contentStart < task.start ||
      task.contentStart > task.end ||
      !record(task.native) ||
      !range(task.native.start, task.native.end, length(task.paragraphId)) ||
      task.native.end - task.native.start !== 1
    )
      return fail('A checklist binding has no native body owner.', 'invalid-document');
    const span = paragraphOffsetIndex(
      owner.node as Extract<OoxmlNode, { kind: 'paragraph' }>
    ).spanOf(control.node);
    if (!span || span.start !== task.native.start || span.end !== task.native.end)
      return fail('A checklist binding has an invalid native range.', 'invalid-document');
    if (tasks.has(task.paragraphId))
      return fail('Duplicate checklist paragraph ownership.', 'ambiguous-target');
    tasks.set(task.paragraphId, task);
  }
  const owners: TargetSourceOwner[] = [],
    byId = new Map<string, TargetSourceOwner>();
  for (const entry of entries.values()) {
    const { node } = entry,
      block = blocks.get(node.id) ?? null;
    const kind =
      node.kind === 'paragraph' && bodyParagraphs.has(node.id)
        ? 'paragraph'
        : node.kind === 'table'
          ? 'table'
          : block
            ? 'block'
            : null;
    if (!kind || entry.outsideStory || !entry.parentId) continue;
    const spans = extents.get(node.id) ?? [];
    const owner: TargetSourceOwner = {
      id: node.id,
      kind,
      parentId: entry.parentId,
      childIndex: entry.childIndex,
      paragraphId: entry.paragraphId,
      cell:
        entry.tableId && entry.rowId && entry.cellId
          ? { tableId: entry.tableId, rowId: entry.rowId, cellId: entry.cellId }
          : null,
      visibility: entry.generated
        ? 'generated'
        : entry.atomic
          ? 'atomic-fallback'
          : spans.length
            ? 'projected'
            : 'unprojected',
      extents: spans,
      sources: slices.get(node.id) ?? [],
      block,
      opaque: opaque.get(node.id) ?? null,
      checklist: tasks.get(node.id) ?? null,
    };
    owners.push(owner);
    byId.set(node.id, owner);
  }
  const windows: TargetSourceWindow[] = [];
  for (const entry of entries.values()) {
    const node = entry.node;
    if (
      entry.outsideStory ||
      entry.generated ||
      entry.atomic ||
      entry.paragraphId !== null ||
      !(node.id === body.id || node.kind === 'tableCell' || node.kind === 'contentControlContent')
    )
      continue;
    const all = children(node);
    let start = -1,
      groups = 0;
    const emit = (end: number) => {
      if (start < 0) return;
      windows.push({
        parentId: node.id,
        startIndex: start,
        ids: all.slice(start, end).map((child) => child.id),
        beforeId: all[start - 1]?.id ?? null,
        afterId: all[end]?.id ?? null,
      });
      groups++;
      start = -1;
    };
    // A blank body paragraph projects to nothing, so Markdown cannot address it. It stays in
    // place as a barrier between windows, like Word's blank lines between paragraphs.
    const blank = (id: string) => {
      const owner = byId.get(id);
      return (
        node.id === body.id &&
        owner?.kind === 'paragraph' &&
        owner.extents.every((extent) => extent.start === extent.end) &&
        paragraphTextOf(part, id) === ''
      );
    };
    for (let i = 0; i < all.length; i++) {
      const owner = byId.get(all[i]!.id),
        eligible =
          owner?.visibility === 'projected' &&
          (owner.kind !== 'block' || owner.paragraphId === null) &&
          !blank(all[i]!.id);
      if (eligible) {
        if (start < 0) start = i;
      } else emit(i);
    }
    emit(all.length);
    if (!groups) {
      let at =
        node.id === body.id && all.length && word(all.at(-1)!, 'sectPr')
          ? all.length - 1
          : all.length;
      // New content goes before the trailing blank paragraphs of an empty body.
      while (at > 0 && blank(all[at - 1]!.id)) at--;
      windows.push({
        parentId: node.id,
        startIndex: at,
        ids: [],
        beforeId: all[at - 1]?.id ?? null,
        afterId: all[at]?.id ?? null,
      });
    }
  }
  const index: TargetSourceIndex = {
    revision: token.plan.revision,
    partName: part.name,
    bodyId: body.id,
    owners,
    windows,
  };
  freeze(index);
  return { success: true, index };
}
