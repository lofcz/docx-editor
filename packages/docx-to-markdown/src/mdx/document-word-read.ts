import {
  buildNumberingIndex,
  buildStyleCascadeTable,
  cascadeParagraphFormatting,
  createListCounterState,
  readNumPr,
  withNumberingStyleLinks,
  type ListCounterAdvance,
} from '@docx-editor.dev/core/layout';
import {
  WML_NAMESPACE_URI as W,
  contentControlPropertiesOf,
  numberingRootOf,
  readOnOffChild,
  sanitizeHref,
  stylesPartOf,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { type DocumentNode } from './document-types';
import { decodeFragment, HeadingFragments } from './document-word-headings';
import { MARKDOWN, MAX_XML_LENGTH, R } from './document-word-package';

type Unwrap = (
  control: OoxmlNode,
  inline: boolean,
  children: () => DocumentNode[]
) => DocumentNode[] | undefined;
type Element = OoxmlElement;
type Branch = DocumentNode & { children: DocumentNode[] };
const branch = (type: string, children: DocumentNode[] = []): Branch => ({ type, children });
const text = (value: string): DocumentNode => ({ type: 'text', value });

function is(node: OoxmlNode, name: string): node is Element {
  return node.kind !== 'textValue' && node.namespaceUri === W && node.localName === name;
}
function child(node: Element | undefined, name: string): Element | undefined {
  for (const item of node?.children ?? []) if (is(item, name)) return item;
  return undefined;
}
function attr(node: Element | undefined, name: string, ns: string = W): string | undefined {
  return node?.attributes.find((a) => a.localName === name && a.namespaceUri === ns)?.value;
}
function val(node: Element | undefined, name: string): string | undefined {
  return attr(child(node, name), 'val');
}
function md(node: Element, name: string): string | undefined {
  return attr(node, name, MARKDOWN);
}
function once<T>(read: () => T): () => T {
  let saved: T | undefined;
  return () => (saved ??= read());
}

const MARKS = new Set(['strong', 'delete', 'emphasis']);
function wrapMarks(node: DocumentNode, order: string[]): DocumentNode {
  return order.reduceRight((item, mark) => branch(mark, [item]), node);
}
function outerMarks(children: DocumentNode[], allowed: readonly string[]) {
  const order: string[] = [];
  while (children.length === 1 && allowed.includes(children[0].type) && children[0].children) {
    order.push(children[0].type);
    children = children[0].children;
  }
  return { order, children };
}
function plain(nodes: DocumentNode[]): string {
  return nodes
    .map((node) => {
      if (node.type === 'break') return '\n';
      if (node.type === 'text' || node.type === 'inlineCode') return node.value ?? '';
      if (MARKS.has(node.type)) return plain(node.children ?? []);
      throw new Error(`Cannot represent ${node.type} inside literal Markdown content`);
    })
    .join('');
}

interface Block {
  nodes: DocumentNode[];
  numbering?: ListCounterAdvance;
  checked?: boolean;
  continuation?: string;
  listSpread?: boolean;
  itemSpread?: boolean;
}

function grouped(blocks: Block[]): DocumentNode[] {
  const result: DocumentNode[] = [];
  const levels: { level: number; id: string; list: Branch; item: Branch }[] = [];
  for (const block of blocks) {
    const number = block.numbering;
    if (block.continuation) {
      const index = levels.map((entry) => entry.id).lastIndexOf(block.continuation);
      if (index >= 0) {
        levels.length = index + 1;
        for (const node of block.nodes) levels[index].item.children.push(node);
        continue;
      }
    }
    if (!number) {
      if (block.nodes.length) levels.length = 0;
      for (const node of block.nodes) result.push(node);
      continue;
    }
    while (levels.length && levels.at(-1)!.level > number.ilvl) levels.pop();
    let active = levels.at(-1);
    if (active?.level !== number.ilvl || active.id !== number.numId) {
      if (active?.level === number.ilvl) levels.pop();
      const ordered = number.level.numFmt !== 'bullet';
      const list = {
        ...branch('list'),
        ordered,
        start: ordered ? number.counters[number.ilvl] : null,
        spread: !!block.listSpread,
      };
      (levels.at(-1)?.item.children ?? result).push(list);
      active = { level: number.ilvl, id: number.numId, list, item: branch('listItem') };
      levels.push(active);
    }
    const item = {
      ...branch('listItem', block.nodes),
      checked: block.checked ?? null,
      spread: !!block.itemSpread,
    };
    active.item = item;
    active.list.children.push(item);
  }
  return result;
}

const INVISIBLE = new Set([
  'bookmarkStart',
  'bookmarkEnd',
  'proofErr',
  'commentRangeStart',
  'commentRangeEnd',
  'commentReference',
  'permStart',
  'permEnd',
  'lastRenderedPageBreak',
  'del',
  'moveFrom',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
]);
const ACCEPTED = new Set(['ins', 'moveTo']);
const alignment = (value: string | undefined) =>
  value === 'right' || value === 'end'
    ? 'right'
    : value === 'left' || value === 'start'
      ? 'left'
      : value === 'center'
        ? 'center'
        : null;

function nameOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') {
    if (!node.value.trim()) return '';
    throw new Error('Unexpected Word text outside a text element');
  }
  if (node.namespaceUri !== W) throw new Error(`Unsupported Word namespace for ${node.localName}`);
  return node.localName;
}

class WordReader {
  private remainingNodes = 1_000_000;
  private remainingText = MAX_XML_LENGTH;
  private readonly counters;
  private readonly cascade;
  private readonly fragments = new HeadingFragments();
  private readonly bookmarks = new Map<string, string>();
  private readonly anchorLinks: { node: DocumentNode; anchor: string; original?: string }[] = [];
  private readonly links = new Map<string, string>();
  private readonly styles = new Map<string, Element>();
  private readonly checkboxes = new WeakMap<DocumentNode, boolean>();
  private readonly cellAlign = new WeakMap<DocumentNode, string | null>();
  private readonly definitions = new Map<string, DocumentNode>();
  private readonly references: { node: DocumentNode; owner: Element }[] = [];

  constructor(
    pkg: OoxmlPackage,
    private readonly unwrap: Unwrap
  ) {
    const styles = stylesPartOf(pkg)?.root;
    this.cascade = buildStyleCascadeTable(styles ?? null);
    this.counters = createListCounterState(
      withNumberingStyleLinks(buildNumberingIndex(numberingRootOf(pkg)), this.cascade)
    );
    for (const rel of pkg.externalTargets) {
      if (rel.type === R + '/hyperlink' && rel.ownerPart === pkg.mainDocumentPart)
        this.links.set(rel.id, rel.rawTarget);
    }
    for (const node of styles?.children ?? []) {
      if (!is(node, 'style')) continue;
      const id = attr(node, 'styleId');
      if (id && !this.styles.has(id)) this.styles.set(id, node);
    }
  }

  private visit(n: OoxmlNode): string {
    if (--this.remainingNodes < 0) throw new Error('Word content exceeds the Markdown node limit');
    return nameOf(n);
  }

  private charge(value: string): void {
    this.remainingText -= value.length;
    if (this.remainingText < 0) throw new Error('Word content exceeds the Markdown text limit');
  }

  private append(nodes: DocumentNode[], next: DocumentNode): void {
    const last = nodes.at(-1);
    if (last?.type === next.type && !this.checkboxes.has(last) && !this.checkboxes.has(next)) {
      if (next.type === 'text' || next.type === 'inlineCode') {
        last.value = (last.value ?? '') + (next.value ?? '');
        return;
      }
      if (MARKS.has(next.type) && last.children && next.children) {
        for (const n of next.children) this.append(last.children, n);
        return;
      }
    }
    nodes.push(next);
  }

  private run(node: Element): DocumentNode[] {
    const props = child(node, 'rPr');
    const style = val(props, 'rStyle');
    const items: DocumentNode[] = [];
    for (const elem of node.children) {
      const name = this.visit(elem);
      if (!name || name === 'rPr' || INVISIBLE.has(name)) continue;
      if (elem.kind === 'textValue') continue;
      switch (name) {
        case 't':
          for (const leaf of elem.children) {
            if (leaf.kind !== 'textValue') throw new Error('Invalid Word text element');
            this.charge(leaf.value);
            if (leaf.value) this.append(items, text(leaf.value));
          }
          break;
        case 'tab':
          this.append(items, text('\t'));
          break;
        case 'noBreakHyphen':
          this.append(items, text('‑'));
          break;
        case 'softHyphen':
          this.append(items, text('\u00ad'));
          break;
        case 'br':
        case 'cr': {
          if (attr(elem, 'type') && attr(elem, 'type') !== 'textWrapping')
            throw new Error('Unsupported Markdown page or column break');
          this.append(items, md(elem, 'soft') === '1' ? text('\n') : { type: 'break' });
          break;
        }
        default:
          throw new Error(`Unsupported Word run content: ${name}`);
      }
    }
    if (!items.length) return [];
    let output = ['HTMLCode', 'CodeChar', 'VerbatimChar'].includes(style ?? '')
      ? [{ type: 'inlineCode', value: plain(items) } as DocumentNode]
      : items;
    const styleProps = child(this.styles.get(style ?? ''), 'rPr');
    const active = new Set<string>();
    for (const [property, mark, named] of [
      ['b', 'strong', 'Strong'],
      ['i', 'emphasis', 'Emphasis'],
      ['strike', 'delete', ''],
    ]) {
      const source = child(props, property) ? props : styleProps;
      if (
        source && child(source, property)
          ? readOnOffChild(source, property)
          : named && named === style
      ) {
        active.add(mark);
      }
    }
    // Run properties decide which marks survive an edit. The optional hint only orders them.
    const hint = (md(node, 'marks') ?? '').split(' ', 3).filter((mark) => active.has(mark));
    const order = [...new Set([...hint, ...active])];
    for (const mark of order.reverse()) output = [branch(mark, output)];
    return output;
  }

  private inline(nodes: readonly OoxmlNode[]): DocumentNode[] {
    const result: DocumentNode[] = [];
    for (const n of nodes) {
      const name = this.visit(n);
      if (!name || name === 'pPr' || INVISIBLE.has(name) || n.kind === 'textValue') continue;
      let items: DocumentNode[];
      if (name === 'r') items = this.run(n);
      else if (ACCEPTED.has(name)) items = this.inline(n.children);
      else if (name === 'sdt') {
        const read = once(() => this.inline(child(n, 'sdtContent')?.children ?? []));
        const custom = this.unwrap(n, true, read);
        if (custom !== undefined) items = custom;
        else {
          const props = contentControlPropertiesOf(n);
          if (props.tag === 'markdown:image') {
            const url = md(n, 'src');
            if (url === undefined) throw new Error('Missing stored Markdown image source');
            this.charge(url);
            const marked = outerMarks(read(), [...MARKS]);
            const label = plain(marked.children);
            const image = {
              type: 'image',
              url,
              title: md(n, 'title') ?? null,
              alt: md(n, 'sourceLabel') && url === label ? '' : label,
            };
            this.reference(n, image);
            items = [wrapMarks(image, marked.order)];
          } else if (props.checkbox) {
            const mark = text(plain(read()));
            this.checkboxes.set(mark, props.checkbox.checked);
            items = [mark];
          } else items = read();
        }
      } else if (name === 'hyperlink') {
        const id = attr(n, 'id', R);
        const anchor = attr(n, 'anchor');
        const raw = id
          ? this.links.get(id)
          : anchor !== undefined
            ? '#' + anchor
            : md(n, 'relative');
        if (raw === undefined) throw new Error('Missing native hyperlink destination');
        const dest = sanitizeHref(raw);
        if (!dest.ok) throw new Error('Unsupported native hyperlink destination');
        this.charge(dest.href);
        const marked = outerMarks(
          this.inline(n.children),
          (md(n, 'outerMarks') ?? '').split(' ', 3).filter((mark) => MARKS.has(mark))
        );
        const link = {
          ...branch('link', marked.children),
          url: dest.href,
          title: attr(n, 'tooltip') ?? null,
        };
        if (!id && anchor !== undefined)
          this.anchorLinks.push({ node: link, anchor, original: md(n, 'fragment') });
        this.reference(n, link);
        items = [wrapMarks(link, marked.order)];
      } else throw new Error(`Unsupported Word inline content: ${name}`);
      for (const item of items) this.append(result, item);
    }
    return result;
  }

  private reference(owner: Element, node: DocumentNode): void {
    if (md(owner, 'identifier') !== undefined) this.references.push({ owner, node });
  }

  root(body: Element): DocumentNode {
    const root = branch('root', grouped(this.blocks(body.children)));
    for (const { node, anchor, original } of this.anchorLinks) {
      const fragment = this.bookmarks.get(anchor);
      if (fragment !== undefined)
        node.url =
          '#' +
          (original !== undefined && decodeFragment(original) === fragment ? original : fragment);
    }
    for (const { owner, node } of this.references) {
      const identifier = md(owner, 'identifier')!;
      const def = this.definitions.get(identifier);
      if (!def || def.url !== node.url || def.title !== node.title) continue;
      delete node.url;
      delete node.title;
      node.type += 'Reference';
      node.identifier = identifier;
      node.label = md(owner, 'label') ?? identifier;
      const type = md(owner, 'referenceType');
      const label =
        node.alt ??
        (node.children?.every((n) => n.type === 'text') ? plain(node.children) : undefined);
      node.referenceType =
        (type === 'shortcut' || type === 'collapsed') && label === node.label ? type : 'full';
    }
    return root;
  }

  private paragraph(n: Element): Block {
    const props = child(n, 'pPr');
    const styleId = val(props, 'pStyle') ?? '';
    const style = this.styles.get(styleId);
    const formatting = cascadeParagraphFormatting(this.cascade, props);
    const num = readNumPr(formatting.paragraphPropertyNodes);
    const numbering = num ? this.counters.advance(num.numId, num.ilvl) : null;
    if (num && !numbering) throw new Error('Unresolved Word numbering level');
    const children = this.inline(n.children);
    let checked: boolean | undefined;
    if (numbering && children[0] && this.checkboxes.has(children[0])) {
      checked = this.checkboxes.get(children[0]);
      children.shift();
      if (children[0]?.type === 'text' && children[0].value?.startsWith(' ')) {
        children[0].value = children[0].value.slice(1);
        if (!children[0].value) children.shift();
      }
    }
    const styleName = val(style, 'name') ?? '';
    const heading = /^heading\s*([1-6])$/i.exec(styleId) ?? /^heading\s*([1-6])$/i.exec(styleName);
    let outline: string | undefined;
    for (const prop of formatting.paragraphProperties) {
      if (prop.localName === 'outlineLvl') outline = prop.attributes?.val;
    }
    const depth = outline !== undefined ? Number(outline) + 1 : heading ? Number(heading[1]) : 0;
    let node: DocumentNode = branch('paragraph', children);
    if (depth >= 1 && depth <= 6 && Number.isInteger(depth)) {
      node = { ...node, type: 'heading', depth };
      const fragment = this.fragments.next(node);
      for (const member of n.children) {
        if (!is(member, 'bookmarkStart') || md(member, 'heading') !== '1') continue;
        const name = attr(member, 'name');
        if (name) this.bookmarks.set(name, fragment);
      }
    } else if (['SourceCode', 'CodeBlock', 'HTMLPreformatted'].includes(styleId)) {
      node = {
        type: 'code',
        value: plain(children),
        lang: md(n, 'lang') ?? null,
        meta: md(n, 'meta') ?? null,
      };
    } else if (!children.length && child(child(props, 'pBdr'), 'bottom')) {
      node = { type: 'thematicBreak' };
    } else if (styleId === 'Quote' || styleId === 'IntenseQuote') {
      node = branch('blockquote', [node]);
    }
    const placeholder = !children.length && (md(n, 'emptyRoot') || md(n, 'listMarker'));
    return {
      nodes: placeholder ? [] : [node],
      numbering: numbering ?? undefined,
      checked,
      listSpread: md(n, 'listSpread') === '1',
      itemSpread: md(n, 'itemSpread') === '1',
    };
  }

  private blocks(nodes: readonly OoxmlNode[]): Block[] {
    const result: Block[] = [];
    for (const n of nodes) {
      const name = this.visit(n);
      if (!name || INVISIBLE.has(name) || n.kind === 'textValue') continue;
      if (['sectPr', 'tblPr', 'tblGrid', 'trPr', 'tcPr'].includes(name)) continue;
      if (name === 'p') result.push(this.paragraph(n));
      else if (name === 'sdt') {
        const blocks = once(() => this.blocks(child(n, 'sdtContent')?.children ?? []));
        const read = once(() => grouped(blocks()));
        const custom = this.unwrap(n, false, read);
        if (custom !== undefined) result.push({ nodes: custom });
        else {
          const tag = contentControlPropertiesOf(n).tag;
          if (tag === 'markdown:quote') result.push({ nodes: [branch('blockquote', read())] });
          else if (tag === 'markdown:continuation')
            result.push({ nodes: read(), continuation: md(n, 'numId') });
          else if (tag === 'markdown:definition') {
            const identifier = md(n, 'identifier');
            const url = md(n, 'url');
            if (!identifier || url === undefined || read().length)
              throw new Error('Invalid Markdown reference definition');
            const definition = {
              type: 'definition',
              identifier,
              url,
              label: md(n, 'label') ?? identifier,
              title: md(n, 'title') ?? null,
            };
            if (!this.definitions.has(identifier)) this.definitions.set(identifier, definition);
            result.push({ nodes: [definition] });
          } else for (const block of blocks()) result.push(block);
        }
      } else if (ACCEPTED.has(name)) {
        for (const block of this.blocks(n.children)) result.push(block);
      } else if (name === 'tbl' || name === 'tr' || name === 'tc') {
        result.push({ nodes: [this.table(n, name)] });
      } else throw new Error(`Unsupported Word block content: ${name}`);
    }
    return result;
  }

  private table(n: Element, name: string): DocumentNode {
    const children = grouped(this.blocks(n.children));
    if (name === 'tc') {
      const props = child(n, 'tcPr');
      if (child(props, 'vMerge') || Number(val(props, 'gridSpan') ?? 1) > 1)
        throw new Error('Merged cells cannot be represented in a Markdown table');
      if (children.length !== 1 || children[0].type !== 'paragraph')
        throw new Error('A Markdown table cell requires one paragraph');
      const cell = branch('tableCell', children[0].children);
      this.cellAlign.set(cell, alignment(val(child(child(n, 'p'), 'pPr'), 'jc')));
      return cell;
    }
    if (name === 'tr') {
      if (!children.length || children.some((cell) => cell.type !== 'tableCell'))
        throw new Error('Invalid Markdown table row content');
      return branch('tableRow', children);
    }
    if (!children.length || children.some((row) => row.type !== 'tableRow'))
      throw new Error('Invalid Markdown table content');
    const align: (string | null)[] = [];
    for (const row of children)
      (row.children ?? []).forEach((cell, i) => {
        if (i >= align.length) align.push(this.cellAlign.get(cell) ?? null);
      });
    return { ...branch('table', children), align };
  }
}

export function readWordDocument(pkg: OoxmlPackage, unwrap: Unwrap): DocumentNode {
  const body = child(pkg.parts.get(pkg.mainDocumentPart)?.root, 'body');
  if (!body) throw new Error('Missing main Word document body');
  return new WordReader(pkg, unwrap).root(body);
}
