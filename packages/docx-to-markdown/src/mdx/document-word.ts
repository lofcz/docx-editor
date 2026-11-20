import { escapeXml, isValidXmlText } from '@docx-editor.dev/core/store';
import { type DocumentNode } from './document-types';
import { decodeFragment } from './document-word-headings';
import {
  MAX_XML_LENGTH,
  WordParts,
  attribute,
  control,
  formatAttribute,
  paragraph,
  style,
  type ListNumbering,
} from './document-word-package';

export { readWordDocument } from './document-word-read';

type Wrap = (node: DocumentNode, inline: boolean, children: () => string) => string | undefined;

const PHRASING = new Set([
  'text',
  'strong',
  'emphasis',
  'delete',
  'link',
  'linkReference',
  'image',
  'imageReference',
  'inlineCode',
  'break',
  'mdxJsxTextElement',
  'mdxTextExpression',
  'inlineMath',
  'footnoteReference',
]);

interface ParagraphContext {
  quote?: number;
  listLevel?: number;
  numbered?: ListNumbering;
  checked?: boolean;
  alignment?: string;
  continuationIndent?: number;
  listAttributes?: string;
}

interface Marks {
  order?: string[];
  link?: boolean;
  code?: boolean;
}

function contents(node: DocumentNode): DocumentNode[] {
  return node.children ?? [];
}

function collect<T>(items: readonly T[], render: (item: T, index: number) => string): string {
  let length = 0;
  const pieces = items.map((item, index) => {
    const piece = render(item, index);
    length += piece.length;
    if (length > MAX_XML_LENGTH) throw new Error('Markdown exceeds the document size limit');
    return piece;
  });
  return pieces.join('');
}

function literal(value: unknown, soft = false): string {
  if (typeof value !== 'string' || !isValidXmlText(value)) throw new Error('Invalid Markdown text');
  if ((value.match(/[\t\n]/g)?.length ?? 0) > 100000)
    throw new Error('Too many Markdown line breaks or tabs');
  return value
    .split(/([\t\n])/)
    .map((part) => {
      if (part === '\t') return '<w:tab/>';
      if (part === '\n') return soft ? '<w:br md:soft="1"/>' : '<w:br/>';
      return part ? `<w:t xml:space="preserve">${escapeXml(part)}</w:t>` : '';
    })
    .join('');
}

function run(content: string, marks: Marks): string {
  const order = marks.order ?? [];
  const properties =
    (marks.code
      ? '<w:rStyle w:val="HTMLCode"/>'
      : marks.link
        ? '<w:rStyle w:val="Hyperlink"/>'
        : '') +
    (order.includes('strong') ? '<w:b/>' : '') +
    (order.includes('emphasis') ? '<w:i/>' : '') +
    (order.includes('delete') ? '<w:strike/>' : '');
  const nesting = order.length > 1 ? formatAttribute('marks', order.join(' ')) : '';
  return `<w:r${nesting}>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}${content}</w:r>`;
}

function checkbox(checked: boolean): string {
  const props =
    `<w14:checkbox><w14:checked w14:val="${checked ? '1' : '0'}"/>` +
    '<w14:checkedState w14:val="2612"/><w14:uncheckedState w14:val="2610"/></w14:checkbox>';
  return control('task', run(literal(checked ? '☒' : '☐'), {}), '', props) + run(literal(' '), {});
}

function pProperties(ctx: ParagraphContext, id?: string, rule = false): string {
  const numbering = ctx.numbered;
  const styleId = id ?? (numbering ? 'ListParagraph' : undefined);
  const borders =
    (ctx.quote ? '<w:left w:val="single" w:sz="8" w:space="4" w:color="888888"/>' : '') +
    (rule ? '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>' : '');
  const indent = (ctx.quote ?? 0) * 360 + (ctx.continuationIndent ?? 0);
  return (
    (styleId ? style(styleId) : '') +
    (numbering
      ? `<w:numPr><w:ilvl w:val="${numbering.level}"/><w:numId w:val="${numbering.id}"/></w:numPr>`
      : '') +
    (borders ? `<w:pBdr>${borders}</w:pBdr>` : '') +
    (indent ? `<w:ind w:left="${indent + (numbering ? (numbering.level + 1) * 720 : 0)}"/>` : '') +
    (ctx.alignment ? `<w:jc w:val="${ctx.alignment}"/>` : '')
  );
}

// The parser supplies an ordinary tree. Bound direct callers, cyclic trees, and oversized scalar values too.
function inspect(root: DocumentNode, parts: WordParts): Map<string, DocumentNode> {
  const headings: DocumentNode[] = [];
  const targets = new Set<string>();
  let nodes = 0;
  let chars = 0;
  const definitions = new Map<string, DocumentNode>();
  function visit(node: DocumentNode, depth: number) {
    if (!node || typeof node.type !== 'string' || depth > 64 || ++nodes > 100000) {
      throw new Error('Invalid or oversized Markdown tree');
    }
    for (const key of ['value', 'url', 'title', 'alt', 'lang', 'meta', 'identifier', 'label']) {
      const field = node[key];
      if (typeof field === 'string') chars += field.length;
    }
    if (chars > MAX_XML_LENGTH || (node.children !== undefined && !Array.isArray(node.children))) {
      throw new Error('Invalid or oversized Markdown content');
    }
    if (
      node.type === 'definition' &&
      typeof node.identifier === 'string' &&
      !definitions.has(node.identifier)
    ) {
      definitions.set(node.identifier, node);
    }
    if (node.type === 'heading') headings.push(node);
    if (
      (node.type === 'link' || node.type === 'definition') &&
      typeof node.url === 'string' &&
      node.url.startsWith('#')
    ) {
      targets.add(decodeFragment(node.url.slice(1)));
    }
    for (const child of contents(node)) visit(child, depth + 1);
  }
  visit(root, 0);
  parts.prepareBookmarks(headings, targets);
  return definitions;
}

class WordWriter {
  constructor(
    readonly parts: WordParts,
    private readonly wrap: Wrap,
    private readonly definitions: Map<string, DocumentNode>
  ) {}

  private wrapped(
    node: DocumentNode,
    inline: boolean,
    children: () => string,
    native: (children: () => string) => string
  ): string {
    let saved: string | undefined;
    const render = () => (saved ??= children());
    const xml = this.wrap(node, inline, render) ?? native(render);
    if (xml.length > MAX_XML_LENGTH) throw new Error('Markdown exceeds the document size limit');
    return xml;
  }

  private inlines(nodes: DocumentNode[], marks: Marks): string {
    return collect(nodes, (node) => this.inline(node, marks));
  }

  private inline(node: DocumentNode, marks: Marks): string {
    const childMarks = ['strong', 'emphasis', 'delete'].includes(node.type)
      ? { ...marks, order: [...new Set([...(marks.order ?? []), node.type])] }
      : node.type === 'link' || node.type === 'linkReference'
        ? { ...marks, link: true }
        : marks;
    const children = () => this.inlines(contents(node), childMarks);
    return this.wrapped(node, true, children, (render) => {
      switch (node.type) {
        case 'text':
          return run(literal(node.value, true), marks);
        case 'break':
          return run('<w:br/>', marks);
        case 'inlineCode':
          return run(literal(node.value), { ...marks, code: true });
        case 'strong':
        case 'emphasis':
        case 'delete':
          return render();
        case 'link':
        case 'linkReference':
        case 'image':
        case 'imageReference': {
          const ref = node.type.endsWith('Reference');
          const source =
            ref && typeof node.identifier === 'string'
              ? this.definitions.get(node.identifier)
              : node;
          if (!source) throw new Error('Missing Markdown reference definition');
          const reference = ref
            ? formatAttribute('identifier', node.identifier) +
              formatAttribute('label', node.label) +
              formatAttribute('referenceType', node.referenceType)
            : '';
          if (node.type.startsWith('link')) {
            const dest = this.parts.hyperlink(source.url);
            const title = source.title == null ? '' : ` w:tooltip="${attribute(source.title)}"`;
            const outer = formatAttribute('outerMarks', marks.order?.join(' '));
            return `<w:hyperlink${dest}${title}${reference}${outer}>${render()}</w:hyperlink>`;
          }
          if (typeof source.url !== 'string') throw new Error('Missing Markdown image source');
          const alt = node.alt ?? '';
          const attrs =
            formatAttribute('src', source.url) +
            formatAttribute('title', source.title) +
            (!alt ? ' md:sourceLabel="1"' : '') +
            reference;
          return control('image', run(literal(alt || source.url), marks), attrs);
        }
        default:
          throw new Error(`Unsupported inline Markdown node: ${node.type}`);
      }
    });
  }

  private flow(nodes: DocumentNode[], ctx: ParagraphContext): string {
    const blocks: (DocumentNode | DocumentNode[])[] = [];
    let pending: DocumentNode[] | undefined;
    for (const node of nodes) {
      if (PHRASING.has(node.type)) {
        if (!pending) blocks.push((pending = []));
        pending.push(node);
      } else {
        pending = undefined;
        blocks.push(node);
      }
    }
    return collect(blocks, (block) =>
      Array.isArray(block)
        ? paragraph(this.inlines(block, {}), pProperties(ctx))
        : this.block(block, ctx)
    );
  }

  block(node: DocumentNode, ctx: ParagraphContext): string {
    const children = () => {
      switch (node.type) {
        case 'paragraph':
        case 'heading':
          return this.inlines(contents(node), {});
        case 'blockquote':
          return this.flow(contents(node), { ...ctx, quote: (ctx.quote ?? 0) + 1 });
        case 'list':
          return this.list(node, ctx);
        case 'table':
          return this.table(node, ctx);
        default:
          return this.flow(contents(node), ctx);
      }
    };
    return this.wrapped(node, false, children, (render) => {
      switch (node.type) {
        case 'root':
          return render() || paragraph('', '', ' md:emptyRoot="1"');
        case 'paragraph':
        case 'heading': {
          let id: string | undefined;
          if (node.type === 'heading') {
            if (
              typeof node.depth !== 'number' ||
              !Number.isInteger(node.depth) ||
              node.depth < 1 ||
              node.depth > 6
            ) {
              throw new Error('Unsupported Markdown heading level');
            }
            id = `Heading${node.depth}`;
          }
          const check = ctx.checked === undefined ? '' : checkbox(ctx.checked);
          return paragraph(
            check + this.parts.headingContent(node, render()),
            pProperties(ctx, id),
            ctx.listAttributes
          );
        }
        case 'thematicBreak':
          return paragraph('', pProperties(ctx, undefined, true));
        case 'code':
          return paragraph(
            run(literal(node.value), {}),
            pProperties(ctx, 'SourceCode'),
            formatAttribute('lang', node.lang) + formatAttribute('meta', node.meta)
          );
        case 'blockquote':
          return control('quote', render());
        case 'list':
        case 'table':
          return render();
        case 'definition': {
          return control(
            'definition',
            '',
            formatAttribute('identifier', node.identifier) +
              formatAttribute('label', node.label) +
              formatAttribute('url', node.url) +
              formatAttribute('title', node.title)
          );
        }
        default:
          throw new Error(`Unsupported block Markdown node: ${node.type}`);
      }
    });
  }

  private list(node: DocumentNode, ctx: ParagraphContext): string {
    if (!contents(node).length) throw new Error('Expected at least one Markdown list item');
    const level = ctx.listLevel ?? 0;
    const numbering = this.parts.numbering(node.ordered === true, node.start ?? 1, level);
    return collect(contents(node), (item) => {
      if (item.type !== 'listItem') throw new Error('Expected a Markdown list item');
      const render = () => {
        const blocks = contents(item);
        const first = blocks[0];
        const listAttributes =
          (node.spread ? ' md:listSpread="1"' : '') + (item.spread ? ' md:itemSpread="1"' : '');
        const numbered = {
          ...ctx,
          numbered: numbering,
          listAttributes,
          checked: typeof item.checked === 'boolean' ? item.checked : undefined,
        };
        const primary = first?.type === 'paragraph' || first?.type === 'heading';
        const mark = primary
          ? this.block(first, numbered)
          : paragraph(
              numbered.checked === undefined ? '' : checkbox(numbered.checked),
              pProperties(numbered),
              ' md:listMarker="1"' + listAttributes
            );
        return (
          mark +
          collect(primary ? blocks.slice(1) : blocks, (block) => {
            if (block.type === 'list') return this.block(block, { ...ctx, listLevel: level + 1 });
            const content = this.block(block, { ...ctx, continuationIndent: (level + 1) * 720 });
            return control('continuation', content, ` md:numId="${numbering.id}"`);
          })
        );
      };
      return this.wrapped(item, false, render, (children) => children());
    });
  }

  private table(node: DocumentNode, ctx: ParagraphContext): string {
    const rows = contents(node);
    const columns = rows.reduce((max, row) => Math.max(max, contents(row).length), 0);
    if (columns < 1 || columns > 63) throw new Error('Unsupported Markdown table column count');
    const align = Array.isArray(node.align) ? node.align : [];
    const width = Math.floor(9360 / columns);
    const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map((edge) => `<w:${edge} w:val="single" w:sz="4" w:color="888888"/>`)
      .join('');
    const properties =
      '<w:tblPr><w:tblStyle w:val="MarkdownTable"/><w:tblW w:w="5000" w:type="pct"/>' +
      `<w:tblBorders>${borders}</w:tblBorders><w:tblLook w:val="0020" w:firstRow="1"/></w:tblPr>`;
    const grid = Array.from({ length: columns }, () => `<w:gridCol w:w="${width}"/>`).join('');
    const body = collect(rows, (row, i) => {
      if (row.type !== 'tableRow') throw new Error('Expected a Markdown table row');
      if (!contents(row).length) throw new Error('Expected at least one Markdown table cell');
      const cells = () =>
        collect(contents(row), (cell, j) => {
          if (cell.type !== 'tableCell') throw new Error('Expected a Markdown table cell');
          const alignment = align[j];
          if (alignment != null && !['left', 'center', 'right'].includes(alignment))
            throw new Error('Invalid Markdown column alignment');
          const children = () =>
            paragraph(
              this.inlines(contents(cell), {}),
              pProperties({ ...ctx, alignment: alignment ?? undefined })
            );
          return this.wrapped(
            cell,
            false,
            children,
            (render) =>
              `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${render()}</w:tc>`
          );
        });
      return this.wrapped(
        row,
        false,
        cells,
        (render) => `<w:tr>${i === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${render()}</w:tr>`
      );
    });
    return `<w:tbl>${properties}<w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
  }
}

export function writeWordDocument(root: DocumentNode, wrap: Wrap): Map<string, Uint8Array> {
  if (root.type !== 'root') throw new Error('Expected a Markdown document root');
  const parts = new WordParts();
  const writer = new WordWriter(parts, wrap, inspect(root, parts));
  return writer.parts.finish(writer.block(root, {}));
}
