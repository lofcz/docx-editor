import {
  applyTreeOp,
  bodyStoryRoot,
  createNodeIdAllocator,
  directParagraphProperties,
  ensureHyperlinkRelationship,
  escapeXml,
  insertChildren,
  readOoxmlPart,
  findNode,
  paragraphTextOf,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlParagraphNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import {
  inlineCodeRunChanges,
  inlineCodeCharacterStyles,
  SCIOBOT_CODE_STYLE_ID,
  withInlineCodeStyle,
} from '../native/inline-code';
import type { NativeEditFailure } from '../contracts/document';
import { DOMAIN_BLOCK_TAGS } from '../projection/block-tags';
import { compileInlineTarget } from './inline-target';
import type { InsertedBlockBinding } from './inserted-blocks';
import type { TargetAstNode } from '../projection/parse-target';
import {
  newParagraphSeeds,
  type NewParagraphCode,
  type NewParagraphLink,
} from '../review/new-paragraphs';
import type { CanonicalEdit } from './plan';

type SiblingEdit = Extract<CanonicalEdit, { kind: 'siblings' }>;

/** A refusal that carries its edit failure code through package-edit completions. */
export class NativeWriteRefusal extends Error {
  constructor(
    message: string,
    readonly code: NativeEditFailure['code'] = 'invalid-input'
  ) {
    super(message);
  }
}
export function refuseWrite(
  message: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): never {
  throw new NativeWriteRefusal(message, code);
}

export const wmlElement = (node: OoxmlNode | undefined | null, name: string): boolean =>
  !!node &&
  node.kind !== 'textValue' &&
  node.namespaceUri === WML_NAMESPACE_URI &&
  node.localName === name;
export function nodeIds(part: OoxmlPart): Set<string> {
  const ids = new Set<string>(),
    stack: OoxmlNode[] = [part.root];
  while (stack.length) {
    const node = stack.pop()!;
    ids.add(node.id);
    if (node.kind !== 'textValue') stack.push(...node.children);
  }
  return ids;
}
const childIds = (part: OoxmlPart, id: string, name: string): string[] => {
  const node = findNode(part, id);
  if (!node || node.kind === 'textValue') return [];
  return node.children.filter((child) => wmlElement(child, name)).map((child) => child.id);
};
export const tableRowIds = (part: OoxmlPart, tableId: string) => childIds(part, tableId, 'tr');
export const rowCellIds = (part: OoxmlPart, rowId: string) => childIds(part, rowId, 'tc');

/** Wrap owned text of paragraphs allocated in this transaction in native hyperlinks. */
export function linkNewParagraphs(
  pkg: OoxmlPackage,
  partName: string,
  paragraphIds: readonly string[],
  links: readonly NewParagraphLink[]
): OoxmlPackage {
  let working = pkg;
  for (const link of links) {
    let target: { relationshipId: string } | { anchor: string };
    if (link.target.kind === 'bookmark') target = { anchor: link.target.anchor };
    else {
      const resource = ensureHyperlinkRelationship(working, link.target.url, partName);
      if (!resource)
        refuseWrite('The native engine refused a new hyperlink target.', 'unsupported-operation');
      working = resource.pkg;
      target = { relationshipId: resource.relationshipId };
    }
    const wrapped = applyTreeOp(working.parts.get(partName)!, {
      op: 'insertHyperlink',
      paragraphId: paragraphIds[link.paragraphIndex]!,
      start: link.start,
      end: link.end,
      ...target,
      ...(link.title !== null ? { tooltip: link.title } : {}),
    });
    if (!wrapped.ok)
      refuseWrite(`A new native hyperlink was refused: ${wrapped.reason}.`, 'apply-failed');
    working = withPart(working, wrapped.part);
  }
  return working;
}

/**
 * Mark code ranges of paragraphs allocated in this transaction with the application character
 * style (N18), adding its definition. Seeds cannot carry `w:rStyle` themselves.
 */
export function codeNewParagraphs(
  pkg: OoxmlPackage,
  partName: string,
  paragraphIds: readonly string[],
  codes: readonly NewParagraphCode[]
): OoxmlPackage {
  if (!codes.length) return pkg;
  const styled = withInlineCodeStyle(pkg);
  if (!styled)
    refuseWrite('The document cannot gain the inline code style.', 'unsupported-operation');
  let working = styled!;
  for (const code of codes) {
    const paragraphId = paragraphIds[code.paragraphIndex]!;
    const changes = inlineCodeRunChanges(
      working.parts.get(partName)!,
      {
        paragraphId,
        start: code.start,
        end: code.end,
        code: true,
      },
      { codeStyles: inlineCodeCharacterStyles(working) }
    );
    if (!changes.success) refuseWrite(changes.error, changes.code);
    for (const change of (changes as Extract<typeof changes, { success: true }>).changes) {
      const styled = withInlineCodeStyle(working, change.after === null ? [] : [change.after]);
      if (!styled)
        refuseWrite('The document cannot gain the inline code style.', 'unsupported-operation');
      working = styled!;
      working = applyOp(working, partName, {
        op: 'setRunProperties',
        paragraphId,
        start: change.start,
        end: change.end,
        properties: change.properties,
        characterStyle: change.after,
      });
    }
  }
  return working;
}

function applyOp(pkg: OoxmlPackage, partName: string, op: TreeDocOp): OoxmlPackage {
  const result = applyTreeOp(pkg.parts.get(partName)!, op);
  if (!result.ok) refuseWrite(`Native ${op.op} refused: ${result.reason}.`, 'apply-failed');
  return withPart(pkg, result.part);
}
/** Merge one direct paragraph property over the authored bag of a new paragraph. */
function setParagraphProperty(
  pkg: OoxmlPackage,
  partName: string,
  paragraphId: string,
  localName: string,
  attributes: Readonly<Record<string, string>>
): OoxmlPackage {
  const part = pkg.parts.get(partName)!,
    properties = directParagraphProperties(part, paragraphId).filter(
      (property) => property.localName !== localName
    );
  return applyOp(pkg, partName, {
    op: 'setParagraphProperties',
    paragraphId,
    properties: [...properties, { localName, attributes: { ...attributes } }],
  });
}

export interface NewCellContent {
  readonly cellId: string;
  readonly blocks: readonly TargetAstNode[];
  /** Word `w:jc` value for every new paragraph in this cell. */
  readonly alignment?: string | null;
}
/** Replace the empty seed paragraph of cells created in this transaction with new paragraphs. */
export function fillNewCells(
  pkg: OoxmlPackage,
  partName: string,
  cells: readonly NewCellContent[]
): OoxmlPackage {
  let working = pkg;
  for (const { cellId, blocks, alignment } of cells) {
    if (!blocks.length) continue;
    const seeds = newParagraphSeeds(blocks, { code: true });
    if (!seeds.success) refuseWrite(seeds.error, seeds.code);
    let part = working.parts.get(partName)!;
    const cell = findNode(part, cellId) as OoxmlElement,
      placeholders = cell.children.filter((node) => node.kind === 'paragraph');
    if (placeholders.length !== 1 || paragraphTextOf(part, placeholders[0]!.id) !== '')
      refuseWrite('A new native cell has no single empty seed paragraph.', 'apply-failed');
    const before = nodeIds(part);
    working = applyOp(working, partName, {
      op: 'insertParagraphs',
      parentId: cellId,
      beforeId: placeholders[0]!.id,
      paragraphs: seeds.paragraphs,
    });
    part = working.parts.get(partName)!;
    const fresh = (findNode(part, cellId) as OoxmlElement).children
      .filter((node) => node.kind === 'paragraph' && !before.has(node.id))
      .map((node) => node.id);
    working = applyOp(working, partName, { op: 'deleteBlock', blockId: placeholders[0]!.id });
    working = linkNewParagraphs(working, partName, fresh, seeds.links);
    working = codeNewParagraphs(working, partName, fresh, seeds.codes);
    if (alignment)
      for (const id of fresh)
        working = setParagraphProperty(working, partName, id, 'jc', { val: alignment });
  }
  return working;
}

interface TableSeed {
  readonly kind: 'table';
  /** Index in `paragraphs` of the new paragraph that follows this table; null uses the gap's `afterId`. */
  readonly anchor: number | null;
  readonly rows: readonly (readonly (readonly TargetAstNode[])[])[];
  readonly alignment: readonly (string | null)[];
}
interface BlockSeed {
  readonly kind: 'block';
  /** Index in `paragraphs` of the new paragraph that follows this block; null uses the gap's `afterId`. */
  readonly anchor: number | null;
  /** Inserted index of the registered element. */
  readonly path: readonly number[];
  readonly ast: TargetAstNode;
}
export interface SiblingFragmentSeeds {
  /** New top-level paragraphs in target order, for one native `insertParagraphs` call. */
  readonly paragraphs: readonly OoxmlParagraphNode[];
  readonly links: readonly NewParagraphLink[];
  readonly codes: readonly NewParagraphCode[];
  /** Paragraph index to resolved paragraph style ID. */
  readonly styles: ReadonlyMap<number, string>;
  /** Top-level inserted index to its paragraph index, for list bindings. */
  readonly paragraphOf: ReadonlyMap<number, number>;
  /** New tables and registered blocks, in target order, placed after allocation. */
  readonly placed: readonly (TableSeed | BlockSeed)[];
}
export type SiblingSeedsResult =
  | { readonly success: true; readonly seeds: SiblingFragmentSeeds }
  | NativeEditFailure;
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
/** Logical targets hold cell paragraphs; raw GFM cells hold inline content directly. */
const cellBlocks = (children: readonly TargetAstNode[]): readonly TargetAstNode[] =>
  !children.length
    ? []
    : children.every((child) => child.type === 'paragraph')
      ? children
      : [{ type: 'paragraph', children: [...children] } as TargetAstNode];
const ALIGNMENT: Readonly<Record<string, string>> = {
  left: 'left',
  center: 'center',
  right: 'right',
};
function tableRows(node: TargetAstNode) {
  const rows = ((node.children ?? []) as TargetAstNode[]).map((row) =>
    row.type !== 'tableRow'
      ? null
      : ((row.children ?? []) as TargetAstNode[]).map((cell) =>
          cell.type === 'tableCell' ? cellBlocks((cell.children ?? []) as TargetAstNode[]) : null
        )
  );
  const width = rows[0]?.length ?? 0;
  if (
    !rows.length ||
    !width ||
    rows.some((row) => !row || row.length !== width || row.some((cell) => cell === null))
  )
    return null;
  const align = Array.isArray(node.align) ? (node.align as unknown[]) : [];
  if (align.some((value) => value !== null && !(String(value) in ALIGNMENT))) return null;
  return {
    rows: rows as TableSeed['rows'],
    alignment: Array.from({ length: width }, (_, column) =>
      align[column] ? ALIGNMENT[String(align[column])]! : null
    ),
  };
}

/**
 * Pure native seeds for one planned sibling insertion: paragraphs, headings, block quotes in the
 * resolved Quote style, GFM tables, and registered application blocks. Shared by direct and
 * suggestion executors; nothing here reads or writes a live editor.
 */
export function siblingFragmentSeeds(
  edit: SiblingEdit,
  /** Whether the gap's `afterId` names a native paragraph a trailing new table can precede. */
  followingParagraph: boolean
): SiblingSeedsResult {
  const styles = new Map(
    (edit.styles ?? []).map((binding) => [binding.path.join('/'), binding.styleId])
  );
  const bound = new Set((edit.blocks ?? []).map((binding) => binding.path.join('/')));
  const flat: TargetAstNode[] = [],
    styleOf = new Map<number, string>(),
    paragraphOf = new Map<number, number>(),
    placed: { before: number; seed: Omit<TableSeed, 'anchor'> | Omit<BlockSeed, 'anchor'> }[] = [];
  // Word needs a paragraph after every table; an empty one projects to nothing.
  const separateTable = () => {
    const last = placed.at(-1);
    if (last?.seed.kind === 'table' && last.before === flat.length)
      flat.push({ type: 'paragraph', children: [] } as TargetAstNode);
  };
  for (const [index, node] of edit.inserted.entries()) {
    if (node.type === 'paragraph' || node.type === 'heading') {
      paragraphOf.set(index, flat.length);
      flat.push(node);
    } else if (node.type === 'blockquote') {
      const style = styles.get(String(index));
      if (!style) return fail('A new block quote needs a resolved native Quote style.');
      const children = (node.children ?? []) as TargetAstNode[];
      if (!children.length || children.some((child) => child.type !== 'paragraph'))
        return fail('New block quotes can contain paragraphs only.');
      for (const child of children) {
        styleOf.set(flat.length, style);
        flat.push(child);
      }
    } else if (node.type === 'table') {
      separateTable();
      const table = tableRows(node);
      if (!table) return fail('A new table needs complete rectangular GFM rows and alignments.');
      placed.push({ before: flat.length, seed: { kind: 'table', ...table } });
    } else if (node.type === 'mdxJsxFlowElement' && bound.has(String(index))) {
      separateTable();
      placed.push({ before: flat.length, seed: { kind: 'block', path: [index], ast: node } });
    } else return fail(`Inserted ${node.type} blocks need their own native representation.`);
  }
  if (!followingParagraph && placed.at(-1)?.seed.kind === 'table') separateTable();
  const seeds = newParagraphSeeds(flat, { code: true });
  if (!seeds.success) return seeds;
  return {
    success: true,
    seeds: {
      paragraphs: seeds.paragraphs,
      links: seeds.links,
      codes: seeds.codes,
      styles: styleOf,
      paragraphOf,
      placed: placed.map(
        ({ before, seed }) =>
          ({ ...seed, anchor: before < flat.length ? before : null }) as TableSeed | BlockSeed
      ),
    },
  };
}

/** Text width of the governing body section, for equal new grid columns. */
function contentWidth(part: OoxmlPart): number {
  const root = bodyStoryRoot(part);
  const section =
    root && root.kind !== 'textValue'
      ? root.children.find((node) => wmlElement(node, 'sectPr'))
      : undefined;
  const value = (name: string, attribute: string) => {
    const node =
      section && section.kind !== 'textValue'
        ? section.children.find((child) => wmlElement(child, name))
        : undefined;
    const raw =
      node && node.kind !== 'textValue'
        ? node.attributes.find((item) => item.localName === attribute)?.value
        : undefined;
    const number = raw === undefined ? NaN : Number(raw);
    return Number.isFinite(number) ? number : null;
  };
  const width = value('pgSz', 'w'),
    left = value('pgMar', 'left') ?? 1440,
    right = value('pgMar', 'right') ?? 1440;
  return width && width - left - right > 720 ? width - left - right : 9360;
}

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MARK_XML: Readonly<Record<string, string>> = {
  bold: '<w:b/>',
  italic: '<w:i/>',
  strike: '<w:strike/>',
};
function runXml(value: string, marks: readonly string[]): string {
  // The code character style (N18) is CT_RPr's first child, ahead of the toggles.
  const properties =
    (marks.includes('code') ? `<w:rStyle w:val="${SCIOBOT_CODE_STYLE_ID}"/>` : '') +
    marks.map((mark) => MARK_XML[mark] ?? '').join('');
  return value
    .split(/(\n|\t|\u001e|\u001f)/)
    .filter(Boolean)
    .map((part) => {
      const content =
        part === '\n'
          ? '<w:br/>'
          : part === '\t'
            ? '<w:tab/>'
            : part === '\u001e'
              ? '<w:noBreakHyphen/>'
              : part === '\u001f'
                ? '<w:softHyphen/>'
                : `<w:t xml:space="preserve">${escapeXml(part)}</w:t>`;
      return `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}${content}</w:r>`;
    })
    .join('');
}

interface RenderContext {
  readonly partName: string;
  readonly width: number;
  readonly bindings: ReadonlyMap<string, InsertedBlockBinding>;
  pkg: OoxmlPackage;
  /** Paragraph AST paths in the order their elements appear in the rendered XML. */
  readonly paragraphs: string[];
}
function paragraphXml(
  node: TargetAstNode,
  path: readonly number[],
  context: RenderContext
): string {
  const inline = compileInlineTarget(node);
  if (!inline.success) refuseWrite(inline.error, inline.code);
  if (
    inline.target.plainText === null ||
    inline.target.tokens.some((token) => token.kind !== 'text')
  )
    refuseWrite(
      'Inline atoms need their own native fragment representation.',
      'unsupported-operation'
    );
  if (inline.target.tokens.some((token) => token.marks.includes('code'))) {
    const styled = withInlineCodeStyle(context.pkg);
    if (!styled)
      refuseWrite('The document cannot gain the inline code style.', 'unsupported-operation');
    context.pkg = styled!;
  }
  const links = new Map(inline.target.hyperlinks.map((link) => [link.path.join('/'), link]));
  let xml = '',
    open: string | null = null;
  for (const token of inline.target.tokens) {
    if (token.kind !== 'text') continue;
    const key = token.hyperlink ? token.hyperlink.join('/') : null;
    if (key !== open) {
      if (open !== null) xml += '</w:hyperlink>';
      open = key;
      if (key !== null) {
        const link = links.get(key)!;
        let target: string;
        if (link.target.kind === 'bookmark') target = `w:anchor="${escapeXml(link.target.anchor)}"`;
        else {
          const made = ensureHyperlinkRelationship(context.pkg, link.target.url, context.partName);
          if (!made)
            refuseWrite(
              'The native engine refused a new hyperlink target.',
              'unsupported-operation'
            );
          context.pkg = made.pkg;
          target = `r:id="${escapeXml(made.relationshipId)}"`;
        }
        xml += `<w:hyperlink ${target}${link.title !== null ? ` w:tooltip="${escapeXml(link.title)}"` : ''}>`;
      }
    }
    xml += runXml(token.value, token.marks);
  }
  if (open !== null) xml += '</w:hyperlink>';
  context.paragraphs.push(path.join('/'));
  const heading =
    node.type === 'heading'
      ? `<w:pPr><w:outlineLvl w:val="${Number(node.depth) - 1}"/></w:pPr>`
      : '';
  return `<w:p>${heading}${xml}</w:p>`;
}
/** A generated paragraph with no target path: empty seeds and atomic fallbacks. */
function seedParagraph(context: RenderContext): string {
  context.paragraphs.push('');
  return '<w:p/>';
}
function blockXml(node: TargetAstNode, path: readonly number[], context: RenderContext): string {
  if (node.type === 'paragraph' || node.type === 'heading')
    return paragraphXml(node, path, context);
  if (node.type === 'table') {
    const table = tableRows(node);
    if (!table)
      refuseWrite('A new table needs complete rectangular GFM rows.', 'unsupported-operation');
    const columns = table.rows[0]!.length,
      width = Math.max(240, Math.floor(context.width / columns));
    const rows = table.rows
      .map(
        (row, r) =>
          `<w:tr>${row
            .map((blocks, c) => {
              const jc = table.alignment[c]
                ? `<w:pPr><w:jc w:val="${table.alignment[c]}"/></w:pPr>`
                : '';
              const content = blocks.length
                ? blocks
                    .map((block, b) =>
                      paragraphXml(block, [...path, r, c, b], context).replace(
                        '<w:p>',
                        `<w:p>${jc}`
                      )
                    )
                    .join('')
                : seedParagraph(context);
              return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
            })
            .join('')}</w:tr>`
      )
      .join('');
    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="' + width + '"/>'.repeat(0)}${Array.from({ length: columns }, () => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
  }
  const binding = context.bindings.get(path.join('/'));
  if (node.type !== 'mdxJsxFlowElement' || !binding)
    refuseWrite(
      `New ${node.type} blocks need their own native representation.`,
      'unsupported-operation'
    );
  const properties = `<w:sdtPr><w:alias w:val="${escapeXml(binding.label)}"/><w:tag w:val="${escapeXml(blockTag(binding.payload.id))}"/></w:sdtPr>`;
  if (DOMAIN_BLOCK_TAGS.get(binding.payload.type)?.kind === 'atomic') {
    const attributes = binding.payload.attributes,
      fallback = ['alt', 'title', 'prompt', 'name', 'url']
        .map((name) => attributes[name])
        .find((value): value is string => typeof value === 'string' && value.length > 0);
    seedParagraph(context);
    return `<w:sdt>${properties}<w:sdtContent><w:p>${runXml(fallback ?? binding.label, [])}</w:p></w:sdtContent></w:sdt>`;
  }
  const children = (node.children ?? []) as TargetAstNode[];
  let content = children.map((child, index) => blockXml(child, [...path, index], context)).join('');
  // A story container needs a paragraph after a final table, and an empty container one seed.
  if (!children.length || children.at(-1)!.type === 'table') content += seedParagraph(context);
  return `<w:sdt>${properties}<w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}
function cloneFresh(node: OoxmlNode, next: () => string): OoxmlNode {
  if (node.kind === 'textValue') return { ...node, id: next() };
  return {
    ...node,
    id: next(),
    children: node.children.map((child) => cloneFresh(child, next)),
  } as OoxmlNode;
}
function paragraphsIn(nodes: readonly OoxmlNode[]): string[] {
  const out: string[] = [];
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') {
      const nested = node.children.some(function has(child: OoxmlNode): boolean {
        return (
          child.kind !== 'textValue' && (child.kind === 'paragraph' || child.children.some(has))
        );
      });
      if (!nested) out.push(node.id);
      return;
    }
    node.children.forEach(visit);
  };
  nodes.forEach(visit);
  return out;
}
/** Render one registered block from target syntax and insert it before `beforeId`, or at the end. */
function placeBlock(
  pkg: OoxmlPackage,
  partName: string,
  parentId: string,
  beforeId: string | null,
  seed: BlockSeed,
  bindings: ReadonlyMap<string, InsertedBlockBinding>
): { pkg: OoxmlPackage; paragraphs: Map<string, string> } {
  let part = pkg.parts.get(partName)!;
  const context: RenderContext = {
    partName,
    width: contentWidth(part),
    bindings,
    pkg,
    paragraphs: [],
  };
  const xml = blockXml(seed.ast, seed.path, context);
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}" xmlns:r="${R}"><w:body>${xml}</w:body></w:document>`,
    { name: '/sciobot-new-block.xml', contentType: part.contentType }
  );
  if (!parsed.ok)
    refuseWrite(`A new application block was refused: ${parsed.reason}.`, 'apply-failed');
  const body = bodyStoryRoot(parsed.part);
  if (!body || body.kind === 'textValue' || body.children.length !== 1)
    refuseWrite('A new application block did not produce one native control.', 'apply-failed');
  pkg = context.pkg;
  part = pkg.parts.get(partName)!;
  const next = createNodeIdAllocator(part, 'new'),
    node = cloneFresh(body.children[0]!, next),
    parent = findNode(part, parentId) as OoxmlElement;
  const terminal = parent.children.findIndex((child) => wmlElement(child, 'sectPr'));
  const at =
    beforeId === null
      ? terminal < 0
        ? parent.children.length
        : terminal
      : parent.children.findIndex((child) => child.id === beforeId);
  if (at < 0)
    refuseWrite('The new block anchor disappeared during the transaction.', 'apply-failed');
  const inserted = insertChildren(part, parentId, at, [node]);
  if (!inserted.ok) refuseWrite('The new application block was refused.', 'apply-failed');
  const ids = paragraphsIn([node]);
  if (ids.length !== context.paragraphs.length)
    refuseWrite('New block paragraphs do not match their target syntax.', 'apply-failed');
  return {
    pkg: withPart(pkg, inserted.part),
    paragraphs: new Map(
      context.paragraphs.flatMap((path, index) => (path ? [[path, ids[index]!] as const] : []))
    ),
  };
}

/**
 * Complete a sibling insertion after native identity allocation: quote styles, links, new tables,
 * and new registered blocks with their payload records. Returns the new paragraph ID of every
 * target paragraph path, for list membership.
 */
export function completeSiblingFragments(
  input: OoxmlPackage,
  partName: string,
  edit: SiblingEdit,
  seeds: SiblingFragmentSeeds,
  fresh: readonly string[]
): { pkg: OoxmlPackage; paragraphs: ReadonlyMap<string, string> } {
  if (fresh.length !== seeds.paragraphs.length)
    refuseWrite('Native sibling allocation did not match the planned gap.', 'apply-failed');
  let pkg = input;
  const paragraphs = new Map<string, string>();
  for (const [index, paragraph] of seeds.paragraphOf)
    paragraphs.set(String(index), fresh[paragraph]!);
  for (const [index, styleId] of seeds.styles)
    pkg = setParagraphProperty(pkg, partName, fresh[index]!, 'pStyle', { val: styleId });
  pkg = linkNewParagraphs(pkg, partName, fresh, seeds.links);
  pkg = codeNewParagraphs(pkg, partName, fresh, seeds.codes);
  const bindings = new Map((edit.blocks ?? []).map((binding) => [binding.path.join('/'), binding]));
  for (const seed of seeds.placed) {
    const anchor = seed.anchor === null ? edit.afterId : fresh[seed.anchor]!;
    if (seed.kind === 'block') {
      const placed = placeBlock(pkg, partName, edit.parentId, anchor, seed, bindings);
      pkg = placed.pkg;
      for (const [path, id] of placed.paragraphs) paragraphs.set(path, id);
      continue;
    }
    const table = seed;
    let part = pkg.parts.get(partName)!;
    if (!anchor || findNode(part, anchor)?.kind !== 'paragraph')
      refuseWrite('A new table needs a following native paragraph.', 'unsupported-operation');
    const columns = table.rows[0]!.length,
      before = new Set(childIds(part, edit.parentId, 'tbl'));
    pkg = applyOp(pkg, partName, {
      op: 'insertTable',
      beforeParagraphId: anchor!,
      rows: table.rows.length,
      cols: columns,
      columnWidthTwips: Math.max(240, Math.floor(contentWidth(part) / columns)),
    });
    part = pkg.parts.get(partName)!;
    const created = childIds(part, edit.parentId, 'tbl').filter((id) => !before.has(id));
    if (created.length !== 1) refuseWrite('Native table allocation is ambiguous.', 'apply-failed');
    const rows = tableRowIds(part, created[0]!);
    if (rows.length !== table.rows.length)
      refuseWrite('The new native table does not match the planned rows.', 'apply-failed');
    pkg = fillNewCells(
      pkg,
      partName,
      rows.flatMap((rowId, row) =>
        rowCellIds(part, rowId).map((cellId, column) => ({
          cellId,
          blocks: table.rows[row]![column]!,
          alignment: table.alignment[column],
        }))
      )
    );
  }
  for (const binding of bindings.values()) {
    const stored = withPayloadRecord(pkg, binding.payload, binding.label);
    if (!stored.ok)
      refuseWrite(`New block payload storage refused: ${stored.reason}.`, 'apply-failed');
    pkg = stored.pkg;
  }
  return { pkg, paragraphs };
}
