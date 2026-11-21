import { firstTableParagraph } from '@docx-editor.dev/core/layout';
import {
  contentControlContentNodeOf,
  contentControlTextOf,
  contentControlPropertiesOf,
  contentControlsIn,
  findCustomXmlDataPart,
  paragraphTextOf,
  readCustomXmlNode,
  readOoxmlPackage,
  withCustomXmlDataPart,
  withCustomXmlNode,
  writeOoxmlPackage,
  writeZip,
  type OoxmlPackage,
  type OoxmlContentControlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import {
  attributesOf,
  attributeSyntax,
  COMPONENT_TAG,
  isComponentSyntax,
  isNativeMark,
  MAX_COMPONENTS,
  METADATA_NS,
  prepareComponentSyntax,
  restoreComponent,
  type WordComponent,
} from './document-components';
import {
  MAX_SOURCE_BYTES,
  parseDocumentSyntax,
  printDocumentSyntax,
  syntaxData,
} from './document-syntax';
import { readWordDocument, writeWordDocument } from './document-word';
import type { JsonValue, MdxComponentRegistry } from './registry';

/** Options for importing a declarative MDX document into Word. @public */
export interface MdxDocumentOptions {
  /** The embedding project's component parsers and Word fallbacks. */
  readonly registry?: MdxComponentRegistry;
  /** Additional block tags with a local expand and collapse control. HTML details is built in. */
  readonly disclosureTags?: readonly string[];
}

/** A component in the main Word story. Expressions in attributes remain inert syntax. @public */
export interface MdxDocumentComponent {
  readonly id: string;
  readonly controlId: string;
  readonly name: string;
  /** Original literal or expression source for math, expressions, and imports. */
  readonly value?: string;
  /** Ordered JSX attributes. Spreads and expression values are retained as inert syntax. */
  readonly jsxAttributes: readonly JsonValue[];
  readonly attributes: Readonly<Record<string, JsonValue>>;
  readonly disclosure: boolean;
  readonly summaryParagraphId: string | null;
}

function open(source: Uint8Array | OoxmlPackage): OoxmlPackage {
  if ('parts' in source) return source;
  const result = readOoxmlPackage(source);
  if (!result.ok) throw new Error('Cannot open the Word document: ' + result.reason + '.');
  return result.package;
}

/**
 * Import Markdown and declarative JSX into editable Word content. Registered components
 * use their Word fallbacks. Other containers retain their editable children. Imports and
 * JavaScript expressions remain source text. This operation never executes them.
 * @public
 */
export function importMdxDocument(source: string, options: MdxDocumentOptions = {}): Uint8Array {
  const root = parseDocumentSyntax(source);
  const { tree, items, wrap } = prepareComponentSyntax(
    root,
    options.registry,
    new Set(['details', ...(options.disclosureTags ?? [])])
  );
  const bytes = writeZip(writeWordDocument(tree, wrap));
  const pkg = open(bytes);
  const store = withCustomXmlDataPart(pkg, pkg.mainDocumentPart, METADATA_NS, 'document');
  if (!store.part) throw new Error('Cannot create the MDX metadata part.');
  const data = JSON.stringify(items);
  if (data.length > MAX_SOURCE_BYTES * 4)
    throw new RangeError('MDX metadata exceeds the document limit.');
  return writeOoxmlPackage(
    withCustomXmlNode(store.pkg, store.part.partName, { id: 'components', label: '', data })
  );
}

/** Return whether a Word document contains the metadata created by importMdxDocument. @public */
export function isMdxDocument(source: Uint8Array | OoxmlPackage): boolean {
  const pkg = open(source);
  return Boolean(findCustomXmlDataPart(pkg, pkg.mainDocumentPart, METADATA_NS));
}

const indexes = new WeakMap<object, ReadonlyMap<string, WordComponent>>();
export function indexedComponents(pkg: OoxmlPackage): ReadonlyMap<string, WordComponent> {
  const part = findCustomXmlDataPart(pkg, pkg.mainDocumentPart, METADATA_NS);
  const tree = part && pkg.parts.get(part.partName);
  if (!tree || !part) return new Map();
  const previous = indexes.get(tree);
  if (previous) return previous;
  const data = readCustomXmlNode(pkg, part.partName, 'components')?.data;
  if (!data || data.length > MAX_SOURCE_BYTES * 4) throw new Error('Invalid MDX metadata part.');
  const values: unknown = JSON.parse(data);
  if (!Array.isArray(values) || values.length > MAX_COMPONENTS)
    throw new Error('Invalid MDX component inventory.');
  const result = new Map<string, WordComponent>();
  for (const raw of values) {
    const row = syntaxData(raw) as unknown as WordComponent;
    if (
      !row ||
      typeof row.id !== 'string' ||
      !/^c[1-9][0-9]{0,7}$/.test(row.id) ||
      result.has(row.id) ||
      !row.node ||
      !isComponentSyntax(row.node) ||
      row.node.children !== undefined ||
      typeof row.disclosure !== 'boolean' ||
      (row.fallbackAttribute !== undefined &&
        (!['alt', 'title', 'label', 'summary', 'name'].includes(row.fallbackAttribute) ||
          row.fallback?.length !== 1 ||
          !attributesOf(row.node).some(
            (entry) => entry.name === row.fallbackAttribute && typeof entry.value === 'string'
          ))) ||
      (row.singleLine !== undefined &&
        (row.singleLine !== true || !row.disclosure || row.node.type !== 'mdxJsxTextElement')) ||
      (row.node.name != null && typeof row.node.name !== 'string') ||
      (row.node.value !== undefined && typeof row.node.value !== 'string') ||
      (row.node.attributes !== undefined && !Array.isArray(row.node.attributes)) ||
      (row.outerMarks !== undefined &&
        (!Array.isArray(row.outerMarks) ||
          row.outerMarks.length > 3 ||
          !row.outerMarks.every(isNativeMark) ||
          new Set(row.outerMarks).size !== row.outerMarks.length)) ||
      (row.caption !== undefined && !['attribute', 'summary', 'name'].includes(row.caption)) ||
      (row.caption === 'summary' && row.summary?.name !== 'summary') ||
      (row.caption === 'attribute' && typeof row.captionAttribute !== 'string') ||
      (row.fallback !== undefined &&
        (!Array.isArray(row.fallback) || row.fallback.some((line) => typeof line !== 'string')))
    )
      throw new Error('Invalid MDX component metadata.');
    result.set(row.id, row);
  }
  indexes.set(tree, result);
  return result;
}

/**
 * Export the current main story to MDX, including edits to native Word paragraphs, lists,
 * tables, and component children. Refuse edits to atomic display text without an MDX mapping.
 * @public
 */
export function exportMdxDocument(source: Uint8Array | OoxmlPackage): string {
  const pkg = open(source);
  const index = indexedComponents(pkg);
  const tree = readWordDocument(pkg, (control, inline, children) => {
    const tag = contentControlPropertiesOf(control).tag;
    if (!tag?.startsWith(COMPONENT_TAG)) return undefined;
    const record = index.get(tag.slice(COMPONENT_TAG.length));
    if (!record) throw new Error('Missing MDX component metadata for ' + tag + '.');
    return [restoreComponent(record, inline, children(), control)];
  });
  return printDocumentSyntax(tree);
}

/** Read component identities and native summary anchors from the current main story. @public */
export function readMdxComponents(
  source: Uint8Array | OoxmlPackage
): readonly MdxDocumentComponent[] {
  const pkg = open(source);
  const part = pkg.parts.get(pkg.mainDocumentPart);
  const index = indexedComponents(pkg);
  if (!part || !index.size) return [];
  const components: MdxDocumentComponent[] = [];
  for (const { node } of contentControlsIn(part.root)) {
    const tag = contentControlPropertiesOf(node).tag;
    if (!tag?.startsWith(COMPONENT_TAG)) continue;
    const record = index.get(tag.slice(COMPONENT_TAG.length));
    if (!record) throw new Error('Missing MDX component metadata for ' + tag + '.');
    components.push(describeComponent(part, node, record, index));
  }
  return components;
}

export function describeComponent(
  part: OoxmlPart,
  node: OoxmlContentControlNode,
  record: WordComponent,
  index: ReadonlyMap<string, WordComponent>
): MdxDocumentComponent {
  const first = contentControlContentNodeOf(node)?.children[0];
  let anchor = first;
  if (record.disclosure && !record.caption) {
    // A block summary owns its original Markdown children. The first flow paragraph anchors the button.
    const tag = first?.kind === 'contentControl' && contentControlPropertiesOf(first).tag;
    if (
      !tag ||
      !tag.startsWith(COMPONENT_TAG) ||
      index.get(tag.slice(COMPONENT_TAG.length))?.node.name !== 'summary'
    )
      anchor = undefined;
    while (anchor?.kind === 'contentControl')
      anchor = contentControlContentNodeOf(anchor)?.children[0];
    if (anchor?.kind === 'table') anchor = firstTableParagraph(anchor);
  }
  const attributes = Object.fromEntries(
    attributesOf(record.node).map(({ name, value }) => [name, value ?? true])
  );
  const display = record.fallbackAttribute ?? record.captionAttribute;
  const value = record.fallbackAttribute
    ? contentControlTextOf(node, '\n')
    : record.caption === 'attribute' && first?.kind === 'paragraph'
      ? (paragraphTextOf(part, first.id) ?? '')
      : undefined;
  if (display && value !== undefined) attributes[display] = value;
  return {
    id: record.id,
    controlId: node.id,
    name: record.node.name ?? record.node.type,
    ...(record.node.value === undefined ? {} : { value: record.node.value }),
    jsxAttributes: attributeSyntax(record.node, display, value),
    attributes: Object.freeze(attributes),
    disclosure: record.disclosure,
    summaryParagraphId: record.disclosure && anchor?.kind === 'paragraph' ? anchor.id : null,
  };
}
