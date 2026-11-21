import {
  contentControlPropertiesOf,
  withCustomXmlDataPart,
  withCustomXmlNode,
  serializeOoxmlPart,
  writeZip,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  COMPONENT_TAG,
  METADATA_NS,
  prepareComponentSyntax,
  restoreComponent,
} from './document-components';
import { MAX_SOURCE_BYTES } from './document-data';
import { indexedComponents, openWordSource } from './document-reader';
import { parseDocumentSyntax, printDocumentSyntax } from './document-syntax';
import { readWordDocument, writeWordDocument } from './document-word';
import type { MdxComponentRegistry } from './registry';

/** Options for importing a declarative MDX document into Word. @public */
export interface MdxDocumentOptions {
  /** The embedding project's component parsers and Word fallbacks. */
  readonly registry?: MdxComponentRegistry;
  /** Additional block tags with a local expand and collapse control. HTML details is built in. */
  readonly disclosureTags?: readonly string[];
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
  const pkg = openWordSource(bytes);
  const store = withCustomXmlDataPart(pkg, pkg.mainDocumentPart, METADATA_NS, 'document');
  if (!store.part) throw new Error('Cannot create the MDX metadata part.');
  const data = JSON.stringify(items);
  if (data.length > MAX_SOURCE_BYTES * 4)
    throw new RangeError('MDX metadata exceeds the document limit.');
  const completed = withCustomXmlNode(store.pkg, store.part.partName, {
    id: 'components',
    label: '',
    data,
  });
  // The Word writer just created and validated these entries. Adding metadata changes
  // only the custom XML store, its relationships, and the package content types.
  const entries = new Map(completed.partBytes);
  const encoder = new TextEncoder();
  for (const [name, part] of completed.parts) {
    if (pkg.parts.get(name) !== part) entries.set(name, encoder.encode(serializeOoxmlPart(part)));
  }
  return writeZip(entries);
}

/**
 * Export the current main story to MDX, including edits to native Word paragraphs, lists,
 * tables, and component children. Refuse edits to atomic display text without an MDX mapping.
 * @public
 */
export function exportMdxDocument(source: Uint8Array | OoxmlPackage): string {
  const pkg = openWordSource(source);
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
