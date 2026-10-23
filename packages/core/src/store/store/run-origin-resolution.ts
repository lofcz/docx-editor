import {
  materializeTrackedSourceBindings,
  readTrackedSourceBindings,
  withTrackedSourceBindings,
} from '../package/tracked-source-bindings.ts';
import { settleDeferredRunOrigins } from './run-origin-deferred.ts';
import {
  INSERTED_PARAGRAPH_PROPERTY_NS,
  hasInsertedPropertyPresence,
} from './inserted-paragraph-property-presence.ts';
import { findNode, parentNodeOf, replaceNode, type EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { PackageEditRefusal, isPackageEditRefusal } from './package-edit-refusal.ts';
import type { TreeOpRejection } from './tree-op-types.ts';
import { joinOriginRuns } from './run-origin-merge.ts';
import { pruneRunOriginNamespaces } from './run-origin-namespaces.ts';
import {
  mapOriginNodes,
  originNodes,
  readRevisionOrigin,
  withoutOriginMetadata,
  isRunOriginNamespace,
  readRunOrigins,
  readTextOrigins,
} from './run-origin-metadata.ts';

interface OriginResolution {
  readonly sourceHints: ReadonlyMap<
    string,
    readonly { readonly scope: string; readonly source: string }[]
  >;
  readonly refusal?: TreeOpRejection;
  readonly paragraphPropertyPrefixes: ReadonlySet<string>;
  readonly paragraphs: ReadonlyMap<string, ReadonlySet<string>>;
  readonly prefixes: ReadonlySet<string>;
  readonly retained: ReadonlySet<string>;
}
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
    if (site.kind !== 'textValue' && hasInsertedPropertyPresence(site))
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
          if (isRunOriginNamespace(attr.namespaceUri) && attr.prefix) prefixes.add(attr.prefix);
        break;
      }
      if (owner.kind !== 'hyperlink') break;
      owner = parentNodeOf(part, owner.id);
    }
  }
  const sourceHints = new Map<string, { scope: string; source: string }[]>();
  try {
    for (const [id, selected] of paragraphs) {
      const paragraph = findNode(part, id);
      if (paragraph)
        sourceHints.set(
          id,
          originNodes(paragraph)
            .flatMap(readRunOrigins)
            .filter((origin) => origin.identity === 'source' && selected.has(origin.scope))
            .map((origin) => ({ scope: origin.scope, source: origin.token.slice(37) }))
        );
    }
  } catch (error) {
    if (!isPackageEditRefusal(error)) throw error;
    return {
      paragraphs,
      prefixes,
      retained,
      paragraphPropertyPrefixes,
      sourceHints,
      refusal: error.reason,
    };
  }
  return { paragraphs, prefixes, retained, paragraphPropertyPrefixes, sourceHints };
}

export function normalizedRunOriginParagraph(
  root: OoxmlNode,
  selected: ReadonlySet<string>,
  prefixes: Set<string>,
  retained: Set<string>,
  hints: readonly { readonly scope: string; readonly source: string }[]
): OoxmlNode {
  root = materializeTrackedSourceBindings(root, selected, false);
  const settled = settleDeferredRunOrigins(root, selected, hints);
  root = settled.root;
  const nodes = originNodes(root),
    scopes = settled.scopes;
  if (!scopes.size) return root;
  for (const node of nodes) {
    if (node.kind === 'textValue') continue;
    for (const origin of [...readRunOrigins(node), ...readTextOrigins(node)]) {
      if (!scopes.has(origin.scope)) continue;
      if ('retainedPrefixes' in origin)
        for (const prefix of origin.retainedPrefixes) retained.add(prefix);
      for (const attr of node.attributes)
        if (attr.namespaceUri === origin.namespaceUri && attr.prefix) prefixes.add(attr.prefix);
    }
  }
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
        readRunOrigins(node).some(
          (run) =>
            scopes.has(run.scope) &&
            readTextOrigins(child).some(
              (text) => text.scope === run.scope && text.namespaceUri === run.namespaceUri
            )
        )
      )
        accessible.add(child.id);
    }
  }
  let changed = false;
  const lexical = mapOriginNodes(nodes, (node) => {
    if (!accessible.has(node.id) || node.kind === 'textValue') return node;
    const textOrigin = readTextOrigins(node).find((origin) => scopes.has(origin.scope));
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
    const origin = readRunOrigins(node).find((origin) => scopes.has(origin.scope));
    const textOrigin = readTextOrigins(node).find((origin) => scopes.has(origin.scope));
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
      let clean = current;
      for (const selectedOrigin of [...readRunOrigins(current), ...readTextOrigins(current)])
        if (scopes.has(selectedOrigin.scope))
          clean = withoutOriginMetadata(clean, selectedOrigin.namespaceUri);
      return clean;
    }
    if (nested.length !== node.children.length) changed = true;
    return current;
  });
  const resolved = changed ? cleaned : root;
  if (resolved.kind !== 'paragraph') return resolved;
  const remaining = readTrackedSourceBindings(resolved);
  return remaining.some((source) => scopes.has(source.scope))
    ? withTrackedSourceBindings(
        resolved,
        remaining.filter((source) => !scopes.has(source.scope))
      )
    : resolved;
}

/** Settle only scopes witnessed at the exact selected native revision sites. */
export function normalizeResolvedRunOrigins(
  part: OoxmlPart,
  resolution: OriginResolution,
  options?: EditOptions
):
  | { readonly ok: true; readonly part: OoxmlPart }
  | { readonly ok: false; readonly reason: TreeOpRejection } {
  if (resolution.refusal) return { ok: false, reason: resolution.refusal };
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
  const prefixes = new Set(resolution.prefixes),
    retained = new Set(resolution.retained);
  for (const [paragraphId, scopes] of resolution.paragraphs) {
    const paragraph = findNode(current, paragraphId);
    if (paragraph?.kind !== 'paragraph') continue;
    const normalized = normalizedRunOriginParagraph(
      paragraph,
      scopes,
      prefixes,
      retained,
      resolution.sourceHints.get(paragraphId) ?? []
    );
    if (normalized === paragraph) continue;
    const made = replaceNode(current, paragraphId, normalized, options);
    if (!made.ok) throw new PackageEditRefusal('Run-origin resolution failed.', 'tree-invariant');
    current = made.part;
  }
  return pruneRunOriginNamespaces(
    current,
    new Set([...prefixes].filter((prefix) => !retained.has(prefix))),
    options
  );
}
