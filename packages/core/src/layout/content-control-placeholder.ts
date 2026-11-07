import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlGenericElementNode,
} from '@docx-editor.dev/core/store';

const placeholders = new WeakSet<OoxmlElement>();
const cache = new WeakMap<OoxmlElement, OoxmlElement>();

/** A local header uses the original paragraph identity, never its text or formatting. */
export function contentControlPlaceholder(source: OoxmlElement): OoxmlElement {
  const known = cache.get(source);
  if (known) return known;
  const property = (name: string, attributes: Record<string, string>): OoxmlGenericElementNode => ({
    kind: 'generic',
    id: `${source.id}:collapse:${name}`,
    namespaceUri: WML_NAMESPACE_URI,
    localName: name,
    prefix: 'w',
    namespaceBindings: [],
    children: [],
    attributes: Object.entries(attributes).map(([localName, value]) => ({
      kind: 'genericExtension',
      namespaceUri: WML_NAMESPACE_URI,
      prefix: 'w',
      localName,
      value,
    })),
  });
  const paragraph: OoxmlElement = {
    id: source.id,
    prefix: 'w',
    namespaceBindings: [],
    attributes: [],
    kind: 'paragraph',
    localName: 'p',
    namespaceUri: WML_NAMESPACE_URI,
    children: [
      {
        kind: 'paragraphProperties',
        id: `${source.id}:collapse:pPr`,
        namespaceUri: WML_NAMESPACE_URI,
        localName: 'pPr',
        prefix: 'w',
        namespaceBindings: [],
        attributes: [],
        children: [
          property('spacing', { before: '0', after: '0', line: '480', lineRule: 'exact' }),
          property('ind', { left: '0', right: '0', firstLine: '0', hanging: '0' }),
          property('pageBreakBefore', { val: '0' }),
          property('keepNext', { val: '0' }),
          property('keepLines', { val: '0' }),
        ],
      },
    ],
  };
  placeholders.add(paragraph);
  cache.set(source, paragraph);
  return paragraph;
}

/** Canonical list counters remain unchanged, but a local header has no list marker. */
export const isContentControlPlaceholder = (paragraph: OoxmlElement): boolean =>
  placeholders.has(paragraph);
