import {
  contentControlContentNodeOf,
  escapeXml,
  isValidXmlText,
  WML_NAMESPACE_URI,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import type {
  BlockPayload,
  JsonObject,
  JsonValue,
  MdxAttribute,
  MdxComponentRegistry,
  TargetAstNode,
} from './registry';
import type { DocumentNode } from './document-types';

export const COMPONENT_TAG = 'mdx:1:';
export const METADATA_NS = 'urn:docx-editor:mdx:v1';
export const MAX_COMPONENTS = 8_192;
export interface WordComponent {
  id: string;
  node: DocumentNode;
  disclosure: boolean;
  caption?: 'attribute' | 'summary' | 'name';
  captionAttribute?: string;
  summary?: DocumentNode;
  fallback?: string[];
  outerMarks?: string[];
}

const WORD_MARKS = new Map([
  ['strong', 'b'],
  ['emphasis', 'i'],
  ['delete', 'strike'],
]);
export const isNativeMark = (value: unknown): value is string =>
  typeof value === 'string' && WORD_MARKS.has(value);
const jsxNode = (node: DocumentNode) =>
  node.type === 'mdxJsxFlowElement' || node.type === 'mdxJsxTextElement';
const literalNodes = new Set([
  'mdxFlowExpression',
  'mdxTextExpression',
  'mdxjsEsm',
  'math',
  'inlineMath',
]);

export function isComponentSyntax(node: DocumentNode): boolean {
  return jsxNode(node) || literalNodes.has(node.type);
}

const paragraph = (value: string): DocumentNode => ({
  type: 'paragraph',
  children: [{ type: 'text', value }],
});
function xmlText(value: string): string {
  if (!isValidXmlText(value)) throw new TypeError('Component text cannot be represented in Word.');
  return escapeXml(value);
}
function run(value: string, marks: readonly string[] = []): string {
  const properties = marks.map((mark) => '<w:' + WORD_MARKS.get(mark) + '/>').join('');
  const order = marks.length > 1 ? ' md:marks="' + marks.join(' ') + '"' : '';
  return (
    '<w:r' +
    order +
    '>' +
    (properties ? '<w:rPr>' + properties + '</w:rPr>' : '') +
    '<w:t xml:space="preserve">' +
    xmlText(value) +
    '</w:t></w:r>'
  );
}
const wordLine = (value: string) => '<w:p>' + run(value) + '</w:p>';

export function attributesOf(node: DocumentNode): MdxAttribute[] {
  const list: MdxAttribute[] = [];
  for (const item of Array.isArray(node.attributes) ? node.attributes : []) {
    if (item?.type === 'mdxJsxAttribute' && typeof item.name === 'string')
      list.push({ name: item.name, value: item.value as JsonValue });
  }
  return list;
}

function withoutChildren(node: DocumentNode): DocumentNode {
  const { children: _children, ...rest } = node;
  return rest;
}

function summaryOf(first: DocumentNode | undefined): DocumentNode | undefined {
  if (first && jsxNode(first) && first.name === 'summary') return first;
  if (first?.type === 'paragraph' && first.children?.length === 1)
    return summaryOf(first.children[0]);
  return undefined;
}

/** Materialize a readable Word fallback for an atomic component with no authored children. */
function fallbackOf(node: DocumentNode, id: string, registry?: MdxComponentRegistry): string[] {
  if (!jsxNode(node)) return [node.value ?? ''];
  const name = node.name ?? '';
  const attrs = attributesOf(node);
  const entry = registry?.parsers.get(name);
  const definition = entry?.atomic ?? entry?.container;
  if (definition?.parse) {
    const attributes = definition.parse.read(attrs, node as unknown as TargetAstNode);
    if (attributes === null) throw new TypeError('The component registry refused ' + name + '.');
    const payload: BlockPayload = { version: 1, id, type: definition.type, attributes };
    if (definition.valid && !definition.valid(payload))
      throw new TypeError('Invalid component payload for ' + name + '.');
    if (definition.wordFallback) return [...definition.wordFallback(payload).lines];
    if (definition.inline) return [definition.inline.fallback(payload).text];
  }
  for (const key of ['alt', 'title', 'label', 'summary', 'name']) {
    const value = attrs.find((attr) => attr.name === key)?.value;
    if (typeof value === 'string' && value) return [value];
  }
  return name ? [name] : [];
}

export function prepareComponentSyntax(
  root: DocumentNode,
  registry: MdxComponentRegistry | undefined,
  disclosures: ReadonlySet<string>
) {
  const items: WordComponent[] = [];
  const owned = new WeakMap<DocumentNode, WordComponent>();
  function visit(source: DocumentNode, outer: string[] = []): DocumentNode {
    if (!isComponentSyntax(source)) {
      const marks =
        isNativeMark(source.type) && !outer.includes(source.type) ? [...outer, source.type] : outer;
      return source.children
        ? { ...source, children: source.children.map((child) => visit(child, marks)) }
        : source;
    }
    if (items.length >= MAX_COMPONENTS)
      throw new RangeError('Too many components in one document.');
    const id = 'c' + (items.length + 1);
    const disclosure = source.type === 'mdxJsxFlowElement' && disclosures.has(source.name ?? '');
    const record: WordComponent = { id, node: withoutChildren(source), disclosure };
    if (outer.length) record.outerMarks = outer;
    items.push(record);
    const children = [...(source.children ?? [])];
    if (disclosure) {
      const summary = source.name === 'details' ? summaryOf(children[0]) : undefined;
      if (summary) {
        record.caption = 'summary';
        record.summary = withoutChildren(summary);
        children[0] = { type: 'paragraph', children: summary.children ?? [] };
      } else {
        const attribute = attributesOf(source).find(
          (attr) =>
            ['title', 'summary', 'label'].includes(attr.name) && typeof attr.value === 'string'
        );
        record.caption = attribute ? 'attribute' : 'name';
        record.captionAttribute = attribute?.name;
        children.unshift(
          paragraph(typeof attribute?.value === 'string' ? attribute.value : (source.name ?? ''))
        );
      }
    } else if ((!jsxNode(source) || children.length === 0) && source.name !== null) {
      record.fallback = fallbackOf(source, id, registry);
      if (record.fallback.some((line) => typeof line !== 'string' || !isValidXmlText(line)))
        throw new TypeError('The component fallback must contain Word-readable text.');
    }
    const node = { ...source, children: children.map((child) => visit(child, outer)) };
    owned.set(node, record);
    return node;
  }
  const tree = visit(root);
  function wrap(node: DocumentNode, inline: boolean, children: () => string): string | undefined {
    const item = owned.get(node);
    if (!item) return undefined;
    const body = item.fallback
      ? inline
        ? run(item.fallback.join(' '), item.outerMarks)
        : item.fallback.map(wordLine).join('')
      : children();
    const alias =
      typeof node.name === 'string' ? '<w:alias w:val="' + xmlText(node.name) + '"/>' : '';
    return (
      '<w:sdt><w:sdtPr>' +
      alias +
      '<w:tag w:val="' +
      COMPONENT_TAG +
      item.id +
      '"/></w:sdtPr><w:sdtContent>' +
      (body || (inline ? run('') : '<w:p/>')) +
      '</w:sdtContent></w:sdt>'
    );
  }
  return { tree, items, wrap };
}

function plainText(node: DocumentNode): string {
  if (node.type === 'break') return '\n';
  return node.children ? node.children.map(plainText).join('') : (node.value ?? '');
}

// A native run can also carry styling that Markdown has no node for, such as underline.
// Check the generated display before reading it through the Markdown projection.
function unchangedDisplayShape(control: OoxmlNode): boolean {
  const content = contentControlContentNodeOf(control);
  function valid(node: OoxmlNode): boolean {
    if (node.kind === 'textValue') return true;
    if (node.namespaceUri !== WML_NAMESPACE_URI) return false;
    if (
      ![
        'p',
        'pPr',
        'r',
        'rPr',
        't',
        'b',
        'i',
        'strike',
        'proofErr',
        'bookmarkStart',
        'bookmarkEnd',
      ].includes(node.localName)
    )
      return false;
    return node.children.every(valid);
  }
  return !!content && content.children.every(valid);
}

export function restoreComponent(
  item: WordComponent,
  inline: boolean,
  content: DocumentNode[],
  control: OoxmlNode
): DocumentNode {
  const node = { ...item.node };
  const outer: string[] = [];
  for (const mark of item.outerMarks ?? []) {
    if (!inline || content.length !== 1 || content[0].type !== mark || !content[0].children) break;
    outer.push(mark);
    content = content[0].children;
  }
  const marked = () =>
    outer.reduceRight<DocumentNode>((child, type) => ({ type, children: [child] }), node);
  if (item.fallback) {
    const unformatted = inline
      ? content.every((child) => child.type === 'text')
      : content.every(
          (child) =>
            child.type === 'paragraph' && child.children?.every((run) => run.type === 'text')
        );
    const printed = inline ? [content.map(plainText).join('')] : content.map(plainText);
    const expected = inline
      ? [item.fallback.join(' ')]
      : item.fallback.length
        ? item.fallback
        : [''];
    if (
      !unformatted ||
      !unchangedDisplayShape(control) ||
      printed.length !== expected.length ||
      printed.some((value, index) => value !== expected[index])
    )
      throw new Error(
        'Cannot map edited display text back to component ' + (node.name ?? node.type) + '.'
      );
    if (jsxNode(node)) node.children = [];
    return marked();
  }
  const children = [...content];
  if (item.caption) {
    const caption = children.shift();
    if (caption?.type !== 'paragraph')
      throw new Error('The disclosure summary paragraph is missing.');
    if (item.caption === 'summary') {
      children.unshift({ ...item.summary!, children: caption.children ?? [] });
    } else {
      const value = plainText(caption);
      if (item.caption === 'attribute' && Array.isArray(node.attributes))
        node.attributes = node.attributes.map((entry: JsonObject) =>
          entry.name === item.captionAttribute ? { ...entry, value } : entry
        );
      else if (value !== node.name)
        throw new Error(
          'A generated component name cannot be changed through its Word display text.'
        );
    }
  }
  node.children = children;
  return marked();
}
