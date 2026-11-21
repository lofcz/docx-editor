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
  type WordComponent,
} from './document-components';
import { MAX_SOURCE_BYTES, syntaxData } from './document-data';
import type { JsonValue } from './registry';

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

export function openWordSource(source: Uint8Array | OoxmlPackage): OoxmlPackage {
  if ('parts' in source) return source;
  const result = readOoxmlPackage(source);
  if (!result.ok) throw new Error('Cannot open the Word document: ' + result.reason + '.');
  return result.package;
}

/** Return whether a Word document contains the metadata created by importMdxDocument. @public */
export function isMdxDocument(source: Uint8Array | OoxmlPackage): boolean {
  const pkg = openWordSource(source);
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

/** Read component identities and native summary anchors from the current main story. @public */
export function readMdxComponents(
  source: Uint8Array | OoxmlPackage
): readonly MdxDocumentComponent[] {
  return componentsIn(openWordSource(source));
}

type DisclosureCache = WeakMap<
  OoxmlContentControlNode,
  { index: ReadonlyMap<string, WordComponent>; component?: MdxDocumentComponent }
>;

/** An ordinary Word edit keeps the identities of untouched controls and the metadata index. */
export function createDisclosureReader(): (pkg: OoxmlPackage) => readonly MdxDocumentComponent[] {
  const cache: DisclosureCache = new WeakMap();
  return (pkg) => componentsIn(pkg, cache);
}

function componentsIn(pkg: OoxmlPackage, disclosures?: DisclosureCache): MdxDocumentComponent[] {
  const part = pkg.parts.get(pkg.mainDocumentPart);
  const index = indexedComponents(pkg);
  if (!part || !index.size) return [];
  const components: MdxDocumentComponent[] = [];
  for (const { node } of contentControlsIn(part.root)) {
    const prior = disclosures?.get(node);
    if (prior?.index === index) {
      if (prior.component) components.push(prior.component);
      continue;
    }
    const tag = contentControlPropertiesOf(node).tag;
    let component: MdxDocumentComponent | undefined;
    if (tag?.startsWith(COMPONENT_TAG)) {
      const record = index.get(tag.slice(COMPONENT_TAG.length));
      if (!record) throw new Error('Missing MDX component metadata for ' + tag + '.');
      if (!disclosures || record.disclosure) {
        component = describeComponent(part, node, record, index);
        components.push(component);
      }
    }
    disclosures?.set(node, { index, component });
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
