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
import { withHorizontalRuleStyle } from './horizontal-rule';

export type TextBlockType = 'p' | 'h1' | 'h2' | 'h3' | 'blockquote';
export const BLOCK_PARAGRAPH_STYLES: Readonly<Record<TextBlockType, string>> = Object.freeze({
  p: 'Normal',
  h1: 'Heading1',
  h2: 'Heading2',
  h3: 'Heading3',
  blockquote: 'Quote',
});

const attribute = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find(
        (entry) => entry.namespaceUri === WML_NAMESPACE_URI && entry.localName === name
      )?.value;
function clone(node: OoxmlNode, next: () => string): OoxmlNode {
  return node.kind === 'textValue'
    ? { ...node, id: next() }
    : ({
        ...node,
        id: next(),
        children: node.children.map((child) => clone(child, next)),
      } as OoxmlElement);
}

/** Add missing named styles before bootstrap. Never replace imported definitions or document defaults. */
export function withBlockTypeStyles(
  pkg: OoxmlPackage,
  options: { readonly preserveConflicts?: boolean } = {}
): OoxmlPackage | null {
  let working = withHorizontalRuleStyle(pkg);
  if (!working) {
    if (!options.preserveConflicts) return null;
    working = pkg;
  }
  const defaults = readOoxmlPackage(blankDocumentBytes());
  if (!defaults.ok) return null;
  const baseline = stylesPartOf(defaults.package)!;
  for (const id of Object.values(BLOCK_PARAGRAPH_STYLES)) {
    const part = stylesPartOf(working);
    if (!part) return options.preserveConflicts ? working : null;
    const matches = part.root.children.filter((node) => attribute(node, 'styleId') === id);
    if (matches.length > 1 || matches.some((node) => attribute(node, 'type') !== 'paragraph')) {
      if (options.preserveConflicts) continue;
      return null;
    }
    if (matches.length) continue;
    const definition = baseline.root.children.find((node) => attribute(node, 'styleId') === id);
    if (!definition || definition.kind !== 'generic' || definition.localName !== 'style')
      return null;
    // Adding a default could change every unstyled paragraph in an imported document.
    const named = {
      ...definition,
      attributes: definition.attributes.filter(
        (entry) => entry.namespaceUri !== WML_NAMESPACE_URI || entry.localName !== 'default'
      ),
    };
    const inserted = insertChildren(part, part.root.id, part.root.children.length, [
      clone(named, createNodeIdAllocator(part)),
    ]);
    if (!inserted.ok) return null;
    working = withPart(working, inserted.part);
  }
  return working;
}
