import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import {
  createNodeIdAllocator,
  insertChildren,
  readOoxmlPackage,
  stylesPartOf,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { withHorizontalRuleStyle } from '../native/horizontal-rule';
import { BLOCK_PARAGRAPH_STYLES } from './block-type';
import { prepareNativeDocument } from '../native/prepare-document';
import { initializedPackageBytes } from '../native/prepared-parts';

const attribute = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)
        ?.value;
function clone(node: OoxmlNode, next: () => string): OoxmlNode {
  return node.kind === 'textValue'
    ? { ...node, id: next() }
    : ({
        ...node,
        id: next(),
        children: node.children.map((child) => clone(child, next)),
      } as OoxmlElement);
}
/** Initialize missing native built-in styles before room bootstrap; never overwrite imported definitions. */
export function withBlockTypeStyles(pkg: OoxmlPackage): OoxmlPackage | null {
  let working = withHorizontalRuleStyle(pkg);
  if (!working) return null;
  const defaults = readOoxmlPackage(blankDocumentBytes());
  if (!defaults.ok) return null;
  const baseline = stylesPartOf(defaults.package)!;
  for (const id of Object.values(BLOCK_PARAGRAPH_STYLES)) {
    const part = stylesPartOf(working)!;
    const matches = part.root.children.filter((node) => attribute(node, 'styleId') === id);
    if (matches.length > 1 || matches.some((node) => attribute(node, 'type') !== 'paragraph'))
      return null;
    if (matches.length) continue;
    const definition = baseline.root.children.find((node) => attribute(node, 'styleId') === id);
    if (!definition) return null;
    const inserted = insertChildren(part, part.root.id, part.root.children.length, [
      clone(definition, createNodeIdAllocator(part)),
    ]);
    if (!inserted.ok) return null;
    working = withPart(working, inserted.part);
  }
  return working;
}

/** Prepare shared infrastructure and block definitions before creating either editor or room. */
export function prepareBlockControlsDocument(document: Uint8Array) {
  const native = prepareNativeDocument(document);
  if (!native.success) return native;
  const loaded = readOoxmlPackage(native.document);
  if (!loaded.ok)
    return {
      success: false as const,
      code: 'invalid-document' as const,
      error: `Cannot read the prepared document: ${loaded.reason}.`,
    };
  const styled = withBlockTypeStyles(loaded.package);
  if (!styled)
    return {
      success: false as const,
      code: 'invalid-document' as const,
      error: 'The document has conflicting block style definitions.',
    };
  return {
    success: true as const,
    document: initializedPackageBytes(native.document, loaded.package, styled),
    changed: native.changed || styled !== loaded.package,
  };
}
