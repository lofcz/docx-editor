// A review decision removes its own paragraph mark, not an unseen peer's future text.
// Keep collaboration identities alive; Word export removes private recovery metadata.
import {
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlParagraphNode,
  type OoxmlPart,
} from './ooxml-tree.ts';
import { WML_NAMESPACE_URI } from './ooxml-shared.ts';
import type { OoxmlPackage } from './ooxml-package.ts';
import { mergeMarkedParagraphs } from '../store/paragraph-mark-merge.ts';

const NS = 'urn:docx-editor.dev:review:retired-paragraph:1';
const isProperties = (node: OoxmlNode) =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'pPr';
export const isRetiredParagraphMarker = (node: OoxmlNode): boolean =>
  node.kind === 'generic' &&
  node.namespaceUri === NS &&
  (node.localName === 'retired' || node.localName === 'joined') &&
  node.children.length === 0 &&
  node.attributes.length === 0;
const NO_HIDDEN: ReadonlySet<string> = new Set();
const BLOCK_CONTAINERS = new Set([
  'body',
  'tc',
  'sdtContent',
  'hdr',
  'ftr',
  'footnote',
  'endnote',
  'txbxContent',
  'customXml',
]);

const NO_STATE = Object.freeze({ retired: false, joined: false });
interface ParagraphState {
  readonly retired: boolean;
  readonly joined: boolean;
}
const stateByParagraph = new WeakMap<OoxmlNode, ParagraphState>();
function paragraphState(node: OoxmlNode): ParagraphState {
  if (node.kind !== 'paragraph') return NO_STATE;
  const cached = stateByParagraph.get(node);
  if (cached) return cached;
  let retired = false,
    joined = false;
  for (const child of node.children) {
    if (!isProperties(child)) continue;
    for (const marker of child.children) {
      if (marker.kind === 'textValue' || !isRetiredParagraphMarker(marker)) continue;
      retired = true;
      joined ||= marker.localName === 'joined';
    }
  }
  const state = retired ? { retired, joined } : NO_STATE;
  stateByParagraph.set(node, state);
  return state;
}
export function isRetiredParagraph(node: OoxmlNode): boolean {
  return paragraphState(node).retired;
}

/** A resolved paragraph mark joins forward in every review display mode. */
export function isJoinedParagraph(node: OoxmlNode): boolean {
  return paragraphState(node).joined;
}

export function isEmptyRetiredParagraph(node: OoxmlNode): boolean {
  const ownsStructure = (child: OoxmlNode): boolean =>
    child.kind !== 'textValue' &&
    ((child.namespaceUri === WML_NAMESPACE_URI &&
      (child.localName === 'sectPr' ||
        child.localName.endsWith('Change') ||
        ['ins', 'del', 'moveFrom', 'moveTo'].includes(child.localName))) ||
      child.children.some(ownsStructure));
  return (
    isRetiredParagraph(node) &&
    node.kind !== 'textValue' &&
    node.children.every(isProperties) &&
    !node.children.some(ownsStructure)
  );
}

/** Never hide the only physical paragraph a container needs. */
export function hiddenRetiredParagraphs(children: readonly OoxmlNode[]): ReadonlySet<string> {
  if (!children.some(isEmptyRetiredParagraph)) return NO_HIDDEN;
  const retired = children.filter(isEmptyRetiredParagraph);
  const normal = children.some(
    (node) => node.kind === 'paragraph' && !isEmptyRetiredParagraph(node)
  );
  return new Set((normal ? retired : retired.slice(1)).map((node) => node.id));
}

export function retireParagraph(
  paragraph: OoxmlParagraphNode,
  mint: () => string,
  forward = false
): OoxmlParagraphNode {
  if (isRetiredParagraph(paragraph) && isJoinedParagraph(paragraph) === forward) return paragraph;
  const marker: OoxmlElement = {
    id: mint(),
    kind: 'generic',
    localName: forward ? 'joined' : 'retired',
    namespaceUri: NS,
    prefix: 'reviewState',
    namespaceBindings: [{ prefix: 'reviewState', namespaceUri: NS }],
    attributes: [],
    children: [],
  };
  const old = paragraph.children.find(isProperties) as OoxmlElement | undefined;
  const properties = (
    old
      ? {
          ...old,
          children: [marker, ...old.children.filter((node) => !isRetiredParagraphMarker(node))],
        }
      : {
          id: mint(),
          kind: 'paragraphProperties',
          namespaceUri: WML_NAMESPACE_URI,
          localName: 'pPr',
          prefix: paragraph.prefix,
          namespaceBindings: [],
          attributes: [],
          children: [marker],
        }
  ) as OoxmlElement;
  return {
    ...paragraph,
    children: [properties, ...paragraph.children.filter((node) => !isProperties(node))],
  } as OoxmlParagraphNode;
}

const exportCache = new WeakMap<OoxmlPart, OoxmlPart>();
const nodeExportCache = new WeakMap<OoxmlNode, OoxmlNode>();
const wordParagraphCache = new WeakMap<OoxmlParagraphNode, OoxmlParagraphNode>();
function wordProperties(properties: OoxmlElement): OoxmlElement {
  let changed = false;
  const children: OoxmlNode[] = [];
  for (const child of properties.children) {
    if (isRetiredParagraphMarker(child)) {
      changed = true;
      continue;
    }
    let next = child;
    if (
      child.kind !== 'textValue' &&
      child.namespaceUri === WML_NAMESPACE_URI &&
      child.localName === 'pPrChange'
    ) {
      const snapshots = child.children.map((snapshot) =>
        isProperties(snapshot) && snapshot.kind !== 'textValue'
          ? wordProperties(snapshot)
          : snapshot
      );
      if (snapshots.some((snapshot, index) => snapshot !== child.children[index]))
        next = { ...child, children: snapshots } as OoxmlElement;
    }
    changed ||= next !== child;
    children.push(next);
  }
  return changed ? ({ ...properties, children } as OoxmlElement) : properties;
}
/** A newly authored paragraph mark cannot inherit an earlier mark's retirement. */
export function withoutRetiredParagraphMetadata(paragraph: OoxmlParagraphNode): OoxmlParagraphNode {
  const cached = wordParagraphCache.get(paragraph);
  if (cached) return cached;
  const children = paragraph.children.map((properties) =>
    isProperties(properties) ? wordProperties(properties) : properties
  );
  const next = children.some((child, index) => child !== paragraph.children[index])
    ? ({ ...paragraph, children } as OoxmlParagraphNode)
    : paragraph;
  wordParagraphCache.set(paragraph, next);
  return next;
}
function exportPart(part: OoxmlPart): OoxmlPart {
  const cached = exportCache.get(part);
  if (cached) return cached;
  const visit = (node: OoxmlNode): OoxmlNode => {
    if (node.kind === 'textValue') return node;
    const cachedNode = nodeExportCache.get(node);
    if (cachedNode) return cachedNode;
    const container =
      node.namespaceUri === WML_NAMESPACE_URI && BLOCK_CONTAINERS.has(node.localName);
    const hidden = container ? hiddenRetiredParagraphs(node.children) : NO_HIDDEN;
    let changed = false;
    const children: OoxmlNode[] = [];
    for (const child of node.children) {
      if (hidden.has(child.id)) {
        changed = true;
        continue;
      }
      // An identical element in opaque payload is not a story paragraph's property marker.
      const stripped =
        container && child.kind === 'paragraph' ? withoutRetiredParagraphMetadata(child) : child;
      const next = visit(stripped);
      changed ||= next !== child;
      children.push(next);
    }
    const joined = container
      ? new Set(node.children.filter(isJoinedParagraph).map((child) => child.id))
      : NO_HIDDEN;
    const merged = joined.size ? mergeMarkedParagraphs(children, joined) : children;
    changed ||= merged !== children;
    const next = changed ? ({ ...node, children: merged } as OoxmlElement) : node;
    nodeExportCache.set(node, next);
    return next;
  };
  const root = visit(part.root) as OoxmlElement;
  const next = root === part.root ? part : { ...part, root };
  exportCache.set(part, next);
  return next;
}

/** A pure Word export snapshot. Live collaboration identities remain in the canonical store. */
export function retiredParagraphExportPackage(pkg: OoxmlPackage): OoxmlPackage {
  let parts: Map<string, OoxmlPart> | null = null;
  for (const [name, part] of pkg.parts) {
    if (
      part.root.namespaceUri !== WML_NAMESPACE_URI ||
      !['document', 'hdr', 'ftr', 'footnotes', 'endnotes'].includes(part.root.localName)
    )
      continue;
    const next = exportPart(part);
    if (next === part) continue;
    parts ??= new Map(pkg.parts);
    parts.set(name, next);
  }
  return parts ? Object.freeze({ ...pkg, parts }) : pkg;
}
