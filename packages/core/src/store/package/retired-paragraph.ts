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
import { rootWithoutProvenanceNamespaces } from '../store/run-origin-namespaces.ts';
import { MAX_XML_DEPTH, XML_DEFAULT_MAX_ELEMENTS } from './xml-reader.ts';
import { PackageEditRefusal } from '../store/package-edit-refusal.ts';

const NS = 'urn:docx-editor.dev:review:retired-paragraph:1';
const FIRST_NS = 'urn:docx-editor.dev:review:retired-paragraph:2';
const isProperties = (node: OoxmlNode) =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'pPr';
export const isRetiredParagraphMarker = (node: OoxmlNode): boolean =>
  node.kind === 'generic' &&
  ((node.namespaceUri === NS && (node.localName === 'retired' || node.localName === 'joined')) ||
    (node.namespaceUri === FIRST_NS && node.localName === 'joined')) &&
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

const NO_STATE = Object.freeze({ retired: false, joined: false, keepFirst: false });
interface ParagraphState {
  readonly retired: boolean;
  readonly joined: boolean;
  readonly keepFirst: boolean;
}
const stateByParagraph = new WeakMap<OoxmlNode, ParagraphState>();
function paragraphState(node: OoxmlNode): ParagraphState {
  if (node.kind !== 'paragraph') return NO_STATE;
  const cached = stateByParagraph.get(node);
  if (cached) return cached;
  let retired = false,
    joined = false,
    keepFirst = false;
  for (const child of node.children) {
    if (!isProperties(child)) continue;
    for (const marker of child.children) {
      if (marker.kind === 'textValue' || !isRetiredParagraphMarker(marker)) continue;
      retired = true;
      joined ||= marker.localName === 'joined';
      keepFirst ||= marker.namespaceUri === FIRST_NS;
    }
  }
  const state = retired ? { retired, joined, keepFirst } : NO_STATE;
  stateByParagraph.set(node, state);
  return state;
}
export function isRetiredParagraph(node: OoxmlNode): boolean {
  return paragraphState(node).retired;
}

/**
 * Whether a resolved paragraph mark joins its text with the following paragraph.
 * The retained canonical paragraph keeps its collaboration identity in every review display mode.
 * Unknown marker versions do not imply a join.
 * @public
 */
export function isJoinedParagraph(node: OoxmlNode): boolean {
  return paragraphState(node).joined;
}

/** Rejected inserted marks retain the first source paragraph's exported identity. */
export function isJoinedParagraphKeepingFirst(node: OoxmlNode): boolean {
  return paragraphState(node).keepFirst;
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

/**
 * Empty resolved paragraphs omitted from the visible children of one native block container.
 * The required final caret paragraph remains visible. Canonical source identities remain intact.
 * @public
 */
export function hiddenRetiredParagraphs(children: readonly OoxmlNode[]): ReadonlySet<string> {
  if (!children.some(isEmptyRetiredParagraph)) return NO_HIDDEN;
  const retired = children.filter(isEmptyRetiredParagraph);
  const normal = children.some(
    (node) => node.kind === 'paragraph' && !isEmptyRetiredParagraph(node)
  );
  if (normal) return new Set(retired.map((node) => node.id));
  let lastTable = -1;
  for (let index = 0; index < children.length; index++)
    if (children[index]!.kind === 'table') lastTable = index;
  // A recovery slot before a table is not the terminal caret paragraph Word requires.
  // Keep a terminal slot when one exists, without adding a leading paragraph to a table-only story.
  const caret =
    lastTable < 0
      ? retired[0]
      : children.find((node, index) => index > lastTable && isEmptyRetiredParagraph(node));
  return new Set(retired.filter((node) => node !== caret).map((node) => node.id));
}

export function retireParagraph(
  paragraph: OoxmlParagraphNode,
  mint: () => string,
  forward = false,
  keepFirst = false
): OoxmlParagraphNode {
  if (
    isRetiredParagraph(paragraph) &&
    isJoinedParagraph(paragraph) === forward &&
    isJoinedParagraphKeepingFirst(paragraph) === (forward && keepFirst)
  )
    return paragraph;
  const namespaceUri = forward && keepFirst ? FIRST_NS : NS;
  const marker: OoxmlElement = {
    id: mint(),
    kind: 'generic',
    localName: forward ? 'joined' : 'retired',
    namespaceUri,
    prefix: 'reviewState',
    namespaceBindings: [{ prefix: 'reviewState', namespaceUri }],
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
const aliasCache = new WeakMap<OoxmlNode, ReadonlyMap<string, ReadonlySet<string>>>();
function retirementAliases(root: OoxmlNode): ReadonlyMap<string, ReadonlySet<string>> {
  let visited = 0;
  const visit = (node: OoxmlNode, depth: number): ReadonlyMap<string, ReadonlySet<string>> => {
    if (++visited > XML_DEFAULT_MAX_ELEMENTS * 2 + 1 || depth > MAX_XML_DEPTH + 1)
      throw new PackageEditRefusal(
        'Retired paragraph export exceeds its traversal limit.',
        'resource-limit'
      );
    const cached = aliasCache.get(node);
    if (cached) return cached;
    const aliases = new Map<string, Set<string>>();
    if (node.kind !== 'textValue') {
      if (isRetiredParagraphMarker(node) && node.prefix)
        aliases.set(node.namespaceUri, new Set([node.prefix]));
      for (const child of node.children)
        for (const [namespaceUri, names] of visit(child, depth + 1)) {
          const prefixes = aliases.get(namespaceUri) ?? new Set<string>();
          for (const prefix of names) prefixes.add(prefix);
          aliases.set(namespaceUri, prefixes);
        }
    }
    aliasCache.set(node, aliases);
    return aliases;
  };
  return visit(root, 0);
}
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
  const aliases = retirementAliases(part.root);
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
      if (hidden.has(child.id) && !isJoinedParagraphKeepingFirst(child)) {
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
    const keepsFirst = new Set(
      node.children.filter(isJoinedParagraphKeepingFirst).map((child) => child.id)
    );
    const merged = joined.size
      ? mergeMarkedParagraphs(children, joined, undefined, undefined, keepsFirst)
      : children;
    changed ||= merged !== children;
    const next = changed ? ({ ...node, children: merged } as OoxmlElement) : node;
    nodeExportCache.set(node, next);
    return next;
  };
  let root = visit(part.root) as OoxmlElement;
  for (const [namespaceUri, prefixes] of aliases)
    root = rootWithoutProvenanceNamespaces(root, prefixes, namespaceUri);
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
