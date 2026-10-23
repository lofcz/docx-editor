import { managedInlineImageAssets } from './inline-image-assets';
import type { ExportSemanticLayout } from '@docx-editor.dev/core/export';
import type { MarkdownLayoutOptions, MarkdownSourceSlice } from '@docx-editor.dev/docx-to-markdown';
import {
  contentControlPropertiesOf,
  paragraphOffsetIndex,
  paragraphTextOf,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import {
  buildBlockIndex,
  type BlockIndexDiagnostic,
  type IndexedBlock,
} from '../blocks/block-index';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { ProjectionComposer } from './compose';
import {
  DOMAIN_BLOCK_TAGS,
  GENERATED_LABEL_TAG,
  attributeLiteral,
  type BlockTagDefinition,
  type BlockPropertyPath,
  type SerializedBlockTag,
} from './block-tags';
import { projectLists, remapChecklists, type ChecklistProjection } from './lists';

export interface BlockAttributeSpan {
  readonly blockId: string;
  readonly controlId: string;
  readonly property: string;
  readonly path: BlockPropertyPath;
  readonly name: string;
  readonly start: number;
  readonly end: number;
  readonly valueStart: number | null;
  readonly valueEnd: number | null;
}
export interface ProjectedBlock {
  readonly id: string;
  readonly controlId: string;
  readonly type: string;
  /** The actual registry policy, including host overrides and custom block types. */
  readonly kind: 'container' | 'atomic';
  readonly tag: string;
  readonly start: number;
  readonly end: number;
  readonly contentStart: number;
  readonly contentEnd: number;
}
export interface BlockProjection extends ChecklistProjection {
  readonly blocks: readonly ProjectedBlock[];
  readonly attributes: readonly BlockAttributeSpan[];
  readonly fallbackParagraphIds: readonly string[];
  readonly blockDiagnostics: readonly BlockIndexDiagnostic[];
}

interface Range {
  start: number;
  end: number;
}
interface Anchor {
  block: IndexedBlock;
  definition: BlockTagDefinition;
  node?: OoxmlNode;
  paragraph?: OoxmlParagraphNode;
  inCell: boolean;
  paragraphIds: Set<string>;
  tableIds: Set<string>;
  labelIds: Set<string>;
  range?: Range;
  labels: Range[];
  children: Anchor[];
}

function sourceMarkdownOffset(
  sources: readonly MarkdownSourceSlice[],
  offset: number,
  edge: 'start' | 'end'
): number | null {
  const found: number[] = [];
  for (const slice of sources) {
    if (slice.sourceScope !== 'body' || offset < slice.sourceStart || offset > slice.sourceEnd)
      continue;
    if (offset === slice.sourceStart) found.push(slice.markdownStart);
    if (offset === slice.sourceEnd) found.push(slice.markdownEnd);
    const boundaries = slice.markdownBoundaries ?? [];
    for (let index = 0; index < boundaries.length; index += 1) {
      const boundary = boundaries[index]!;
      if (boundary.sourceOffset === offset) found.push(boundary.markdownOffset);
      const next = boundaries[index + 1];
      if (
        next &&
        boundary.sourceOffset < offset &&
        offset < next.sourceOffset &&
        next.sourceOffset - boundary.sourceOffset === next.markdownOffset - boundary.markdownOffset
      )
        found.push(boundary.markdownOffset + offset - boundary.sourceOffset);
    }
  }
  return found.length
    ? found.reduce((best, value) =>
        edge === 'start' ? Math.max(best, value) : Math.min(best, value)
      )
    : null;
}

/** Wrap read-only native constructs. Payload properties and native child text keep separate targets. */
export function projectBlocks(
  pkg: OoxmlPackage,
  layout: ExportSemanticLayout,
  revision: RevisionToken,
  definitions: ReadonlyMap<string, BlockTagDefinition> = DOMAIN_BLOCK_TAGS,
  options: Pick<MarkdownLayoutOptions, 'tableSyntax'> = {}
): { readonly success: true; readonly projection: BlockProjection } | NativeEditFailure {
  const index = buildBlockIndex(pkg);
  const anchors = new Map<string, Anchor>();
  for (const block of index.byId.values()) {
    const definition = definitions.get(block.payload.type);
    if (definition && block.partName === pkg.mainDocumentPart)
      anchors.set(block.controlId, {
        block,
        definition,
        inCell: false,
        paragraphIds: new Set(),
        tableIds: new Set(),
        labelIds: new Set(),
        labels: [],
        children: [],
      });
  }
  const part = pkg.parts.get(pkg.mainDocumentPart);
  if (!part)
    return { success: false, code: 'invalid-document', error: 'The document has no body part.' };
  let visits = 0;
  const visit = (
    node: OoxmlNode,
    parents: readonly Anchor[],
    paragraph: OoxmlParagraphNode | undefined,
    inCell: boolean,
    labelOwner: Anchor | undefined,
    depth: number
  ): void => {
    if (++visits > 100_000 || depth > 64) throw new Error('Block traversal limit exceeded.');
    if (node.kind === 'textValue') return;
    let active = parents;
    const anchor = anchors.get(node.id);
    if (anchor) {
      if (labelOwner) throw new Error('Generated labels cannot contain registered blocks.');
      anchor.node = node;
      anchor.paragraph = paragraph;
      anchor.inCell = inCell;
      active = [...parents, anchor];
    }
    if (
      node.kind === 'contentControl' &&
      contentControlPropertiesOf(node).tag === GENERATED_LABEL_TAG
    )
      labelOwner = active.at(-1);
    if (node.kind === 'paragraph') {
      paragraph = node;
      for (const owner of active) owner.paragraphIds.add(node.id);
      if (labelOwner && active.at(-1) === labelOwner) labelOwner.labelIds.add(node.id);
    }
    if (node.kind === 'table') {
      if (labelOwner) throw new Error('Generated labels cannot contain tables.');
      for (const owner of active) owner.tableIds.add(node.id);
    }
    if (node.kind === 'tableCell') inCell = true;
    for (const child of node.children)
      visit(child, active, paragraph, inCell, labelOwner, depth + 1);
  };
  try {
    visit(part.root, [], undefined, false, undefined, 0);
    // Rich opaque trees need cell paragraph boundaries that GFM cannot express.
    const richCells = [...anchors.values()].some(
      (anchor) =>
        anchor.node &&
        anchor.inCell &&
        anchor.definition.serialize &&
        anchor.definition.serializedCellSyntax !== 'gfm'
    );
    if (richCells && options.tableSyntax === 'gfm')
      return {
        success: false,
        code: 'unsupported-operation',
        error: 'Opaque cell content needs HTML table syntax.',
      };
    const listed = projectLists(pkg, layout, revision, {
      tableSyntax: richCells ? 'html' : (options.tableSyntax ?? 'gfm'),
      imageAssets: managedInlineImageAssets(
        pkg,
        layout,
        [...anchors.values()].map((anchor) => anchor.block)
      ),
    });
    if (!listed.success) return listed;
    const base = listed.projection;
    const paragraphConstructs = new Map(
      base.sourceMap.paragraphs
        .filter((value) => value.sourceScope === 'body')
        .map((value) => [value.paragraphId, value])
    );
    const tableConstructs = new Map(
      base.sourceMap.tables
        .filter((value) => value.sourceScope === 'body')
        .map((value) => [value.tableId, value])
    );
    const paragraphSources = new Map<string, MarkdownSourceSlice[]>();
    for (const source of base.sourceMap.sources) {
      if (source.sourceScope !== 'body') continue;
      const values = paragraphSources.get(source.paragraphId) ?? [];
      values.push(source);
      paragraphSources.set(source.paragraphId, values);
    }
    for (const anchor of anchors.values()) {
      if (!anchor.node) continue;
      const constructs: Range[] = [];
      if (anchor.paragraph) {
        const span = paragraphOffsetIndex(anchor.paragraph).spanOf(anchor.node);
        if (!span) throw new Error('An inline block has no native source span.');
        const wholeParagraph =
          span.start === 0 && span.end === paragraphTextOf(part, anchor.paragraph.id)?.length;
        const paragraph = paragraphConstructs.get(anchor.paragraph.id);
        // Inline ownership covers its native text, not generated heading/list/mark syntax.
        const start =
          sourceMarkdownOffset(
            paragraphSources.get(anchor.paragraph.id) ?? [],
            span.start,
            'start'
          ) ?? (wholeParagraph ? paragraph?.markdownStart : undefined);
        const end =
          sourceMarkdownOffset(paragraphSources.get(anchor.paragraph.id) ?? [], span.end, 'end') ??
          (wholeParagraph ? paragraph?.markdownEnd : undefined);
        if (start !== undefined && start !== null && end !== undefined && end !== null)
          constructs.push({ start, end });
      } else {
        for (const id of anchor.paragraphIds) {
          const paragraph = paragraphConstructs.get(id);
          if (
            !paragraph ||
            paragraph.markdownStart === undefined ||
            paragraph.markdownEnd === undefined
          )
            continue;
          const range = { start: paragraph.markdownStart, end: paragraph.markdownEnd };
          constructs.push(range);
          if (anchor.labelIds.has(paragraph.paragraphId)) anchor.labels.push(range);
        }
        for (const id of anchor.tableIds) {
          const table = tableConstructs.get(id);
          if (table) constructs.push({ start: table.markdownStart, end: table.markdownEnd });
        }
      }
      if (constructs.length)
        anchor.range = {
          start: constructs.reduce((start, value) => Math.min(start, value.start), Infinity),
          end: constructs.reduce((end, value) => Math.max(end, value.end), -Infinity),
        };
    }
    const visible = [...anchors.values()].filter((anchor) => anchor.range);
    const roots: Anchor[] = [];
    const byId = new Map(visible.map((anchor) => [anchor.block.id, anchor]));
    for (const anchor of visible) {
      let parentId = anchor.block.parentId;
      let parent: Anchor | undefined;
      while (parentId) {
        parent = byId.get(parentId);
        if (parent) break;
        parentId = index.byId.get(parentId)?.parentId ?? null;
      }
      (parent ? parent.children : roots).push(anchor);
    }
    const writer = new ProjectionComposer(base);
    const blocks: ProjectedBlock[] = [];
    const attributes: BlockAttributeSpan[] = [];
    const fallbackParagraphIds: string[] = [];
    const fallbackTableIds = new Set<string>();
    function tag(anchor: Anchor, atomic: boolean): SerializedBlockTag {
      if (anchor.definition.serialize) {
        if (!atomic) throw new Error('Only atomic blocks can serialize complete stored syntax.');
        const serialized = anchor.definition.serialize(anchor.block.payload);
        if (typeof serialized.text !== 'string' || !serialized.text.length)
          throw new Error('Serialized block syntax is empty.');
        for (const property of serialized.properties) {
          if (
            !Number.isInteger(property.start) ||
            !Number.isInteger(property.end) ||
            property.start < 0 ||
            property.end < property.start ||
            property.end > serialized.text.length ||
            (property.valueStart === null) !== (property.valueEnd === null) ||
            (property.valueStart !== null &&
              (!Number.isInteger(property.valueStart) ||
                !Number.isInteger(property.valueEnd) ||
                property.valueStart < property.start ||
                property.valueEnd! < property.valueStart ||
                property.valueEnd! > property.end))
          )
            throw new Error('Serialized block property has invalid bounds.');
        }
        return serialized;
      }
      if (!/^[A-Za-z][A-Za-z0-9_.:-]*$/.test(anchor.definition.tag))
        throw new Error('Invalid registered tag name.');
      let text = `<${anchor.definition.tag}`;
      const attributes: Omit<BlockAttributeSpan, 'blockId' | 'controlId'>[] = [];
      const names = new Set<string>();
      for (const binding of anchor.definition.attributes(anchor.block.payload)) {
        if (!/^[A-Za-z_][A-Za-z0-9_.:-]*$/.test(binding.name) || names.has(binding.name))
          throw new Error('Invalid or duplicate block attribute.');
        names.add(binding.name);
        const value = attributeLiteral(binding.value);
        if (value === null) continue;
        text += ' ';
        const start = text.length;
        text += binding.name;
        let valueStart: number | null = null,
          valueEnd: number | null = null;
        if (binding.value !== true) {
          text += '="';
          valueStart = text.length;
          text += value;
          valueEnd = text.length;
          text += '"';
        }
        attributes.push({
          name: binding.name,
          property: binding.property,
          path: Object.freeze(['attributes', binding.property]),
          start,
          end: text.length,
          valueStart,
          valueEnd,
        });
      }
      text += atomic ? ' />' : '>';
      return { tag: anchor.definition.tag, text, properties: attributes };
    }
    function renderRange(
      start: number,
      end: number,
      children: readonly Anchor[],
      labels: readonly Range[] = []
    ): void {
      const pieces = [
        ...children.map((anchor) => ({ range: anchor.range!, anchor })),
        ...labels.map((range) => ({ range, anchor: undefined })),
      ].sort((a, b) => a.range.start - b.range.start || b.range.end - a.range.end);
      let cursor = start;
      for (const piece of pieces) {
        if (piece.range.start < cursor || piece.range.end > end)
          throw new Error('Registered block source ranges overlap.');
        if (cursor < piece.range.start) writer.native(cursor, piece.range.start);
        if (piece.anchor) render(piece.anchor);
        cursor = piece.range.end;
        if (!piece.anchor && base.text.slice(cursor, cursor + 2) === '\n\n' && cursor + 2 <= end)
          cursor += 2;
      }
      if (cursor < end || (start === end && pieces.length === 0)) writer.native(cursor, end);
    }
    function render(anchor: Anchor): void {
      const atomic = anchor.definition.kind === 'atomic';
      const opening = tag(anchor, atomic);
      const start = writer.offset;
      for (const span of opening.properties)
        attributes.push(
          Object.freeze({
            ...span,
            path: Object.freeze([...span.path]),
            blockId: anchor.block.id,
            controlId: anchor.block.controlId,
            start: start + span.start,
            end: start + span.end,
            valueStart: span.valueStart === null ? null : start + span.valueStart,
            valueEnd: span.valueEnd === null ? null : start + span.valueEnd,
          })
        );
      let contentStart: number, contentEnd: number;
      if (atomic) {
        writer.replacement(anchor.range!.start, anchor.range!.end, opening.text);
        contentStart = contentEnd = writer.offset;
        fallbackParagraphIds.push(...anchor.paragraphIds);
        for (const id of anchor.tableIds) fallbackTableIds.add(id);
      } else {
        const inline = Boolean(anchor.paragraph) || anchor.inCell;
        writer.generated(opening.text + (inline ? '' : '\n\n'));
        contentStart = writer.offset;
        renderRange(anchor.range!.start, anchor.range!.end, anchor.children, anchor.labels);
        contentEnd = writer.offset;
        writer.generated((inline ? '' : '\n\n') + `</${anchor.definition.tag}>`);
      }
      fallbackParagraphIds.push(...anchor.labelIds);
      blocks.push(
        Object.freeze({
          id: anchor.block.id,
          controlId: anchor.block.controlId,
          type: anchor.block.payload.type,
          kind: anchor.definition.kind,
          tag: opening.tag,
          start,
          end: writer.offset,
          contentStart,
          contentEnd,
        })
      );
    }
    renderRange(0, base.text.length, roots);
    const projection = writer.finish();
    const hidden = new Set(fallbackParagraphIds);
    return {
      success: true,
      projection: Object.freeze({
        ...projection,
        sourceMap: Object.freeze({
          ...projection.sourceMap,
          paragraphs: Object.freeze(
            projection.sourceMap.paragraphs.filter(
              (paragraph) => !hidden.has(paragraph.paragraphId)
            )
          ),
          tables: Object.freeze(
            projection.sourceMap.tables.filter((table) => !fallbackTableIds.has(table.tableId))
          ),
        }),
        blocks: Object.freeze(blocks.sort((a, b) => a.start - b.start || b.end - a.end)),
        checklists: remapChecklists(writer, base.checklists, hidden),
        attributes: Object.freeze(attributes),
        fallbackParagraphIds: Object.freeze([...new Set(fallbackParagraphIds)]),
        blockDiagnostics: Object.freeze(
          index.diagnostics.map((diagnostic) => Object.freeze({ ...diagnostic }))
        ),
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

/** Attribute values address application payload properties, never native fallback text. */
export function blockAttributeAt(
  projection: BlockProjection,
  offset: number
): BlockAttributeSpan | null {
  if (!Number.isInteger(offset)) return null;
  return projection.attributes.find((span) => span.start <= offset && offset < span.end) ?? null;
}
