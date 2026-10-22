import {
  INSERTED_PARAGRAPH_PROPERTY_NS,
  preservesInsertedParagraphProperty,
} from './inserted-paragraph-property-presence.ts';
import {
  findNode,
  parentNodeOf,
  replaceChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { PackageEditRefusal, isPackageEditRefusal } from './package-edit-refusal.ts';
import type { TreeOpRejection } from './tree-op-types.ts';
import { joinOriginRuns } from './run-origin-merge.ts';
import { pruneRunOriginNamespaces } from './run-origin-namespaces.ts';
import {
  children,
  mapOriginNodes,
  originNodes,
  readRevisionOriginScope,
  readRevisionOrigin,
  readRunOrigin,
  readTextOrigin,
  withoutOriginMetadata,
  RUN_ORIGIN_NS,
} from './run-origin-metadata.ts';

interface OriginResolution {
  readonly paragraphPropertyPrefixes: ReadonlySet<string>;
  readonly paragraphs: ReadonlyMap<string, ReadonlySet<string>>;
  readonly prefixes: ReadonlySet<string>;
  readonly retained: ReadonlySet<string>;
}
const revision = (node: OoxmlNode) =>
  node.kind !== 'textValue' &&
  (node.kind.startsWith('revision') || /PrChange$/.test(node.localName));

/** Capture before the native rebuild transfers the warm node index to the resulting root. */
export function captureRunOriginResolution(
  part: OoxmlPart,
  sites: readonly OoxmlNode[]
): OriginResolution {
  const paragraphs = new Map<string, Set<string>>();
  const prefixes = new Set<string>();
  const retained = new Set<string>();
  const paragraphPropertyPrefixes = new Set<string>();
  for (const site of sites) {
    if (preservesInsertedParagraphProperty(site, 'pPr') && site.kind !== 'textValue')
      for (const attr of site.attributes)
        if (attr.namespaceUri === INSERTED_PARAGRAPH_PROPERTY_NS && attr.prefix)
          paragraphPropertyPrefixes.add(attr.prefix);
    const origin = readRevisionOrigin(site);
    if (!origin) continue;
    const { scope } = origin;
    let owner = parentNodeOf(part, site.id);
    if (site.kind !== 'textValue' && site.localName === 'rPrChange') {
      if (owner?.kind !== 'runProperties') continue;
      owner = parentNodeOf(part, owner.id);
      if (owner?.kind !== 'run') continue;
      owner = parentNodeOf(part, owner.id);
    } else if (site.kind !== 'revisionInsert' && site.kind !== 'revisionDelete') continue;
    for (let depth = 0; owner && depth <= 256; depth++) {
      if (owner.kind === 'paragraph') {
        const scopes = paragraphs.get(owner.id) ?? new Set<string>();
        scopes.add(scope);
        paragraphs.set(owner.id, scopes);
        for (const prefix of origin.retainedPrefixes) retained.add(prefix);
        for (const attr of site.attributes)
          if (attr.namespaceUri === RUN_ORIGIN_NS && attr.prefix) prefixes.add(attr.prefix);
        break;
      }
      if (owner.kind !== 'hyperlink') break;
      owner = parentNodeOf(part, owner.id);
    }
  }
  return { paragraphs, prefixes, retained, paragraphPropertyPrefixes };
}

function normalizedParagraph(root: OoxmlNode, selected: ReadonlySet<string>): OoxmlNode {
  const nodes = originNodes(root);
  const scopes = new Set(selected);
  for (const node of nodes) {
    if (!revision(node)) continue;
    const scope = readRevisionOriginScope(node);
    if (scope) scopes.delete(scope);
  }
  if (!scopes.size) return root;
  // Only ordinary runs reached through the paragraph and hyperlinks participate.
  // Controls, drawings, foreign revisions, and property snapshots remain opaque.
  const accessible = new Set<string>([root.id]);
  for (const node of nodes) {
    if (!accessible.has(node.id) || node.kind === 'textValue') continue;
    for (const child of node.children) {
      if (
        (node.kind === 'paragraph' || node.kind === 'hyperlink') &&
        (child.kind === 'run' || child.kind === 'hyperlink')
      )
        accessible.add(child.id);
      if (node.kind === 'run' && child.kind === 'runProperties') accessible.add(child.id);
      if (
        node.kind === 'run' &&
        (child.kind === 'text' || child.kind === 'deletedText') &&
        readRunOrigin(node)?.scope === readTextOrigin(child)?.scope &&
        scopes.has(readRunOrigin(node)?.scope ?? '')
      )
        accessible.add(child.id);
    }
  }
  let changed = false;
  const lexical = mapOriginNodes(nodes, (node) => {
    if (!accessible.has(node.id) || node.kind === 'textValue') return node;
    const textOrigin = readTextOrigin(node);
    if (!textOrigin || !scopes.has(textOrigin.scope)) return node;
    const attributes = node.attributes.filter(
      (attr) =>
        !(
          attr.namespaceUri === 'http://www.w3.org/XML/1998/namespace' && attr.localName === 'space'
        )
    );
    if (textOrigin.space !== 'absent')
      attributes.push({
        kind: 'xmlSpace',
        namespaceUri: 'http://www.w3.org/XML/1998/namespace',
        localName: 'space',
        prefix: 'xml',
        value: textOrigin.space as 'default' | 'preserve',
      });
    changed = true;
    return { ...node, attributes } as OoxmlNode;
  });
  const joined = mapOriginNodes(originNodes(lexical), (node) => {
    if (!accessible.has(node.id) || (node.kind !== 'paragraph' && node.kind !== 'hyperlink'))
      return node;
    const nested: OoxmlNode[] = [];
    for (const child of node.children) {
      const previous = nested.at(-1);
      const merged = previous && joinOriginRuns(previous, child, scopes);
      if (merged) {
        nested[nested.length - 1] = merged;
        changed = true;
      } else nested.push(child);
    }
    return nested.length === node.children.length
      ? node
      : ({ ...node, children: nested } as OoxmlNode);
  });
  const joinedNodes = originNodes(joined);
  const cleaned = mapOriginNodes(joinedNodes, (node) => {
    if (node.kind === 'textValue' || !accessible.has(node.id)) return node;
    const origin = readRunOrigin(node);
    const textOrigin = readTextOrigin(node);
    const selected =
      (origin && scopes.has(origin.scope)) || (textOrigin && scopes.has(textOrigin.scope));
    const nested = node.children.filter(
      (child) =>
        !(
          node.kind === 'run' &&
          selected &&
          origin?.absent &&
          child.kind === 'runProperties' &&
          !child.children.length &&
          !child.attributes.length
        )
    );
    const current =
      nested.length === node.children.length
        ? node
        : ({ ...node, children: nested } as typeof node);
    if (selected) {
      changed = true;
      return withoutOriginMetadata(current);
    }
    if (nested.length !== node.children.length) changed = true;
    return current;
  });
  return changed ? cleaned : root;
}

/** Settle only scopes witnessed at the exact selected native revision sites. */
export function normalizeResolvedRunOrigins(
  part: OoxmlPart,
  resolution: OriginResolution,
  options?: EditOptions
):
  | { readonly ok: true; readonly part: OoxmlPart }
  | { readonly ok: false; readonly reason: TreeOpRejection } {
  try {
    return {
      ok: true,
      part: pruneRunOriginNamespaces(
        normalizeOrigins(part, resolution, options),
        resolution.paragraphPropertyPrefixes,
        options,
        INSERTED_PARAGRAPH_PROPERTY_NS
      ),
    };
  } catch (error) {
    if (isPackageEditRefusal(error)) return { ok: false, reason: error.reason };
    throw error;
  }
}

function normalizeOrigins(
  part: OoxmlPart,
  resolution: OriginResolution,
  options?: EditOptions
): OoxmlPart {
  let current = part;
  for (const [paragraphId, scopes] of resolution.paragraphs) {
    const paragraph = findNode(current, paragraphId);
    if (paragraph?.kind !== 'paragraph') continue;
    const normalized = normalizedParagraph(paragraph, scopes);
    if (normalized === paragraph) continue;
    const made = replaceChildren(current, paragraphId, children(normalized), options);
    if (!made.ok) throw new PackageEditRefusal('Run-origin resolution failed.', 'tree-invariant');
    current = made.part;
  }
  return pruneRunOriginNamespaces(
    current,
    new Set([...resolution.prefixes].filter((prefix) => !resolution.retained.has(prefix))),
    options
  );
}
