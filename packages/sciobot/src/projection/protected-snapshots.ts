import { sha256FontBytes } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

export type ProtectionReason =
  | 'opaque-xml'
  | 'field'
  | 'drawing'
  | 'inline-atom'
  | 'range-marker'
  | 'section-properties'
  | 'outside-body'
  | 'element-shell';
export interface ProtectedNodeReference {
  readonly partName: string;
  readonly nodeId: string;
  readonly parentId: string | null;
  readonly kind: ProtectionReason;
  readonly reason: string;
  readonly mode: 'subtree' | 'shell';
  readonly fingerprint: string;
}
export interface ProtectedPartReference {
  readonly partName: string;
  readonly mode: 'xml' | 'bytes';
  readonly fingerprint: string;
}
export interface ProtectedParentOrder {
  readonly partName: string;
  readonly nodeId: string;
  readonly childIds: readonly string[];
}
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const XML = 'http://www.w3.org/XML/1998/namespace';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const nativePropertyAttributes = new Set(
  (
    'val ascii hAnsi eastAsia cs asciiTheme hAnsiTheme eastAsiaTheme cstheme hint ' +
    'color fill themeColor themeTint themeShade themeFill themeFillTint themeFillShade ' +
    'before after beforeLines afterLines beforeAutospacing afterAutospacing line lineRule ' +
    'left right start end firstLine hanging leftChars rightChars startChars endChars firstLineChars hangingChars ' +
    'sz space shadow frame pos leader count sep equalWidth type w h x y origin orient bidi'
  ).split(/\s+/)
);

/** Attributes outside the native body's ordinary text ownership remain protected. */
function unknownAttribute(
  node: Exclude<OoxmlNode, { kind: 'textValue' }>,
  namespace: string,
  name: string
): boolean {
  if (namespace === XML && name === 'space') return false;
  if (namespace === W14 && (name === 'paraId' || name === 'textId')) return false;
  if (namespace === W && /^rsid/.test(name)) return false;
  if (
    node.kind === 'hyperlink' &&
    ((namespace === R && name === 'id') ||
      (namespace === W &&
        ['anchor', 'history', 'tooltip', 'docLocation', 'tgtFrame'].includes(name)))
  )
    return false;
  if (
    ['revisionInsert', 'revisionDelete', 'revisionMoveFrom', 'revisionMoveTo'].includes(
      node.kind
    ) &&
    namespace === W &&
    ['id', 'author', 'date'].includes(name)
  )
    return false;
  if (node.kind === 'generic' && namespace === W && nativePropertyAttributes.has(name))
    return false;
  return true;
}

export const bytesFingerprint = (bytes: Uint8Array): string => sha256FontBytes(bytes);
export const dataFingerprint = (value: unknown): string =>
  bytesFingerprint(new TextEncoder().encode(JSON.stringify(value)));
export function nodeFingerprint(node: OoxmlNode, mode: 'subtree' | 'shell'): string {
  if (mode === 'subtree') {
    const bindings: unknown[] = [],
      stack = [node];
    while (stack.length) {
      const next = stack.pop()!;
      if (next.kind === 'textValue') continue;
      bindings.push(
        [...next.namespaceBindings].sort((a, b) =>
          JSON.stringify(a).localeCompare(JSON.stringify(b))
        )
      );
      for (let at = next.children.length - 1; at >= 0; at--) stack.push(next.children[at]!);
    }
    return dataFingerprint([canonicalOoxmlFingerprint(node), bindings]);
  }
  if (node.kind === 'textValue') return dataFingerprint(null);
  return dataFingerprint({
    namespaceUri: node.namespaceUri,
    localName: node.localName,
    attributes: node.attributes
      .filter((attribute) => unknownAttribute(node, attribute.namespaceUri, attribute.localName))
      .map((attribute) => [attribute.namespaceUri, attribute.localName, attribute.value])
      .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
    bindings: [...node.namespaceBindings].sort((a, b) =>
      JSON.stringify(a).localeCompare(JSON.stringify(b))
    ),
  });
}
export function hasProtectedShell(node: OoxmlNode): boolean {
  return (
    node.kind !== 'textValue' &&
    (node.namespaceBindings.length > 0 ||
      node.attributes.some((attribute) =>
        unknownAttribute(node, attribute.namespaceUri, attribute.localName)
      ))
  );
}
export function packageMetadataFingerprint(pkg: OoxmlPackage): string {
  const sorted = <T>(entries: Iterable<readonly [string, T]>) =>
    [...entries].sort(([a], [b]) => a.localeCompare(b));
  return dataFingerprint({
    mainDocumentPart: pkg.mainDocumentPart,
    relationships: sorted(pkg.relationships),
    defaults: sorted(pkg.contentTypes.defaults),
    overrides: sorted(pkg.contentTypes.overrides),
    externalTargets: pkg.externalTargets,
  });
}
export function protectedPartsOf(pkg: OoxmlPackage): readonly ProtectedPartReference[] {
  const parts: ProtectedPartReference[] = [];
  for (const [partName, part] of pkg.parts)
    if (partName !== pkg.mainDocumentPart)
      parts.push(
        Object.freeze({
          partName,
          mode: 'xml',
          fingerprint: nodeFingerprint(part.root, 'subtree'),
        })
      );
  for (const [partName, bytes] of pkg.partBytes)
    if (partName !== pkg.mainDocumentPart && !pkg.parts.has(partName))
      parts.push(Object.freeze({ partName, mode: 'bytes', fingerprint: bytesFingerprint(bytes) }));
  return Object.freeze(parts.sort((a, b) => a.partName.localeCompare(b.partName)));
}
