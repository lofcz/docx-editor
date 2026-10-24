import { secureRandomUuid } from '../package/random-uuid.ts';
import {
  withTrackedSourceBindings,
  readTrackedSourceBindings,
  type TrackedSourceBinding,
} from '../package/tracked-source-bindings.ts';
import { sourceOriginIdentity } from './run-origin-identity.ts';
import {
  createNodeIdAllocator,
  findNode,
  parentNodeOf,
  replaceChildren,
  replaceNode,
} from '../package/ooxml-edit.ts';
import { namespaceBindingsAt, wmlFreshNamespaceContextAt } from '../package/wml-namespace.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';
import { build } from './tree-op-tracked-builders.ts';
import { collectRevisionSites } from './tree-op-revisions.ts';
import { effectiveContentLockAt, isBoundAt } from './tree-op-nodes.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';
import { sourceRunShells, inheritSourceRunShell } from './run-origin-shells.ts';
import {
  children,
  encodeRetainedPrefixes,
  mapOriginNodes,
  originNodes,
  originPrefix,
  readRunOrigin,
  runProperties,
  withOriginMetadata,
  withoutOriginMetadata,
  originNamespaces,
  readRunOrigins,
  readTextOrigins,
  readRevisionOrigins,
  isRunOriginNamespace,
  originNamespaceForScope,
} from './run-origin-metadata.ts';

/** Native source-run provenance prepared within the actual package transaction. @public */
export interface TrackedRunOriginPreparation {
  readonly part: OoxmlPart;
  readonly scope: string | null;
}

function ordinaryRun(part: OoxmlPart, id: string, paragraphId: string): boolean {
  let node = findNode(part, id);
  if (node?.kind !== 'run') return false;
  for (let depth = 0; node && depth <= 256; depth++) {
    if (node.id === paragraphId) return true;
    if (node.kind !== 'run' && node.kind !== 'hyperlink') return false;
    node = parentNodeOf(part, node.id);
  }
  return false;
}

/**
 * Retain original run boundaries for later tracked text/link composition without copying text.
 * Only runs touching the supplied original-coordinate edits receive provenance.
 * Refuses protected/opaque source interiors and active or opaque source provenance.
 * Recognized inactive source provenance remains unchanged.
 * @public
 */
export function prepareTrackedRunOrigins(
  part: OoxmlPart,
  paragraphId: string,
  ranges: readonly { readonly start: number; readonly end: number }[]
): TrackedRunOriginPreparation {
  const paragraph = findNode(part, paragraphId);
  if (paragraph?.kind !== 'paragraph')
    throw new PackageEditRefusal('The run-origin paragraph is missing.', 'unknown-paragraph');
  if (isBoundAt(part, paragraphId) || effectiveContentLockAt(part, paragraphId).content)
    throw new PackageEditRefusal('The run-origin paragraph is protected.', 'unsupported');
  if (!Array.isArray(ranges) || ranges.length > 4096)
    throw new PackageEditRefusal('Run-origin ranges exceed their limit.', 'resource-limit');
  const index = paragraphOffsetIndex(paragraph);
  for (const range of ranges) {
    if (
      !range ||
      !Number.isSafeInteger(range.start) ||
      !Number.isSafeInteger(range.end) ||
      range.start < 0 ||
      range.end < range.start ||
      range.end > index.length
    )
      throw new PackageEditRefusal('Invalid run-origin range.', 'invalidArgs');
  }
  const ordered = [...ranges].sort((a, b) => a.start - b.start);
  const merged: { start: number; end: number }[] = [];
  for (const range of ordered) {
    const last = merged.at(-1);
    if (last && range.start <= last.end) last.end = Math.max(last.end, range.end);
    else merged.push({ start: range.start, end: range.end });
  }
  const ids = new Set<string>();
  let cursor = 0;
  for (const segment of index.segments) {
    while (merged[cursor] && merged[cursor]!.end < segment.start) cursor++;
    const range = merged[cursor];
    if (!range || range.start > segment.end) continue;
    let interior = false;
    for (let at = cursor; merged[at] && merged[at]!.start <= segment.end; at++) {
      const candidate = merged[at]!;
      interior ||=
        candidate.start === candidate.end
          ? segment.start < candidate.start && segment.end > candidate.start
          : segment.start < candidate.end && segment.end > candidate.start;
    }
    const eligible =
      !segment.removeNodeIds &&
      !segment.formatRunIds &&
      ordinaryRun(part, segment.runId, paragraphId);
    if (!eligible) {
      if (interior)
        throw new PackageEditRefusal('Run-origin source is protected or opaque.', 'unsupported');
      continue;
    }
    ids.add(segment.runId);
  }
  if (!ids.size) return { part, scope: null };
  const nodes = originNodes(paragraph);
  const activeOrigins = nodes
    .filter(
      (node) =>
        node.kind === 'revisionInsert' ||
        node.kind === 'revisionDelete' ||
        runPropertySite(part, node)
    )
    .flatMap(readRevisionOrigins);
  const active = new Set(activeOrigins.map((origin) => origin.scope));
  const activeNamespaces = new Set(activeOrigins.map((origin) => origin.namespaceUri));
  const preparedBindings = readTrackedSourceBindings(paragraph);
  for (const run of nodes) {
    if (!ids.has(run.id)) continue;
    const origins = readRunOrigins(run);
    const properties = runProperties(run);
    const provisional = origins.length ? [] : readRunOrigins(properties);
    // Repeated ordered stages may touch an already prepared source fragment.
    // Only registered, unsealed proofs may be layered; active and opaque proofs still refuse.
    const reusable =
      provisional.length > 0 &&
      provisional.every((origin) =>
        preparedBindings.some(
          (binding) =>
            binding.scope === origin.scope &&
            binding.token === origin.token &&
            origin.identity === 'source' &&
            !origin.resolved &&
            origin.absent === binding.absent &&
            JSON.stringify(origin.retainedPrefixes) === JSON.stringify(binding.retained)
        )
      );
    const proofs = reusable ? provisional : origins;
    for (const node of originNodes(run)) {
      const namespaces = originNamespaces(node);
      if (!namespaces.length) continue;
      const valid =
        node === run
          ? origins
          : reusable && node === properties
            ? provisional
            : children(run).some((child) => child === node) &&
                (node.kind === 'text' || node.kind === 'deletedText')
              ? readTextOrigins(node).filter(
                  (text) =>
                    proofs.some(
                      (origin) =>
                        origin.scope === text.scope && origin.namespaceUri === text.namespaceUri
                    ) &&
                    (!reusable ||
                      preparedBindings.some(
                        (binding) =>
                          binding.scope === text.scope &&
                          binding.texts.some(
                            (original) =>
                              original.token === text.token && original.space === text.space
                          )
                      ))
                )
              : [];
      if (valid.length !== namespaces.length || valid.some((origin) => active.has(origin.scope)))
        throw new PackageEditRefusal(
          'Selected source already carries active or opaque run-origin metadata.',
          'unsupported'
        );
    }
  }
  const inherited = namespaceBindingsAt(part, paragraph);
  const retained = encodeRetainedPrefixes([
    ...new Set([
      ...[...inherited]
        .filter(([, uri]) => isRunOriginNamespace(uri) && !activeNamespaces.has(uri))
        .map(([prefix]) => prefix),
      ...nodes.flatMap((node) =>
        node.kind === 'textValue'
          ? []
          : node.namespaceBindings
              .filter(
                (binding) =>
                  isRunOriginNamespace(binding.namespaceUri) &&
                  !activeNamespaces.has(binding.namespaceUri)
              )
              .map((binding) => binding.prefix)
      ),
    ]),
  ]);
  const sourceIdentity = sourceOriginIdentity(nodes);
  const scope = secureRandomUuid();
  const prefix = originPrefix(nodes, inherited, `runOrigin${scope.replace(/-/g, '')}`);
  const mint = createNodeIdAllocator(part);
  const sourceBindings: TrackedSourceBinding[] = [];
  const next = mapOriginNodes(nodes, (node) => {
    if (node.kind !== 'run' || !ids.has(node.id)) return node;
    const old = runProperties(node);
    let properties: OoxmlElement;
    if (old) properties = old;
    else {
      const context = wmlFreshNamespaceContextAt(part, node);
      properties = {
        ...build(mint(), 'runProperties', 'rPr', [], []),
        prefix: context.elementPrefix,
        namespaceBindings: context.rowBinding ? [context.rowBinding] : [],
      } as OoxmlElement;
    }
    sourceBindings.push({
      scope,
      runId: node.id,
      token: `${scope}/${sourceIdentity(node)}`,
      absent: !old,
      retained: retained ? (JSON.parse(retained) as string[]) : [],
      texts: node.children
        .filter((child) => child.kind === 'text')
        .map((child) => ({
          id: child.id,
          token: `${scope}/${sourceIdentity(child)}`,
          space:
            child.attributes.find(
              (attr) =>
                attr.namespaceUri === 'http://www.w3.org/XML/1998/namespace' &&
                attr.localName === 'space'
            )?.value ?? 'absent',
        })),
    });
    const marked = withOriginMetadata(
      properties,
      {
        origin: `${scope}/${sourceIdentity(node)}`,
        identity: 'source',
        ...(!old ? { absent: '1' } : {}),
        ...(retained ? { retained } : {}),
      },
      prefix
    );
    return {
      ...node,
      children: [
        ...(!old ? [marked] : []),
        ...node.children.map((child) => {
          if (child.id === old?.id) return marked;
          if (child.kind !== 'text') return child;
          const space =
            child.attributes.find(
              (attr) =>
                attr.namespaceUri === 'http://www.w3.org/XML/1998/namespace' &&
                attr.localName === 'space'
            )?.value ?? 'absent';
          if (!['absent', 'default', 'preserve'].includes(space))
            throw new PackageEditRefusal(
              'The source whitespace setting is unsupported.',
              'unsupported'
            );
          return withOriginMetadata(
            child,
            { textOrigin: `${scope}/${sourceIdentity(child)}`, space, identity: 'source' },
            prefix
          );
        }),
      ],
    } as OoxmlNode;
  });
  if (next.kind !== 'paragraph') throw new Error('Invalid run-origin paragraph.');
  const bound = withTrackedSourceBindings(
    next,
    [...readTrackedSourceBindings(paragraph), ...sourceBindings],
    originPrefix(nodes, inherited, 'sourceBindings')
  );
  const made = replaceNode(part, paragraphId, bound);
  if (!made.ok) throw new PackageEditRefusal('Run-origin annotation failed.', 'tree-invariant');
  return { part: made.part, scope };
}

const address = (node: OoxmlNode) =>
  node.kind === 'textValue'
    ? null
    : JSON.stringify(
        ['id', 'author', 'date'].map(
          (name) =>
            node.attributes.find(
              (attr) => attr.namespaceUri === WML_NAMESPACE_URI && attr.localName === name
            )?.value ?? null
        )
      );

function runPropertySite(part: OoxmlPart, node: OoxmlNode): boolean {
  if (node.kind === 'textValue' || node.localName !== 'rPrChange') return false;
  let owner = parentNodeOf(part, node.id);
  for (let depth = 0; owner && depth <= 256; depth++) {
    if (owner.kind === 'run') return true;
    if (owner.kind === 'paragraph') return false;
    owner = parentNodeOf(part, owner.id);
  }
  return false;
}

/**
 * Attach provenance only to newly allocated native revision sites in the prepared paragraph.
 * Prior site identities and complete Word addresses remain excluded, including copied histories.
 * @public
 */
export function sealTrackedRunOrigins(
  part: OoxmlPart,
  paragraphId: string,
  scope: string,
  before: OoxmlPart
): OoxmlPart {
  if (part.name !== before.name)
    throw new PackageEditRefusal('Run-origin stories do not match.', 'invalidArgs');
  const paragraph = findNode(part, paragraphId);
  if (paragraph?.kind !== 'paragraph')
    throw new PackageEditRefusal('The run-origin paragraph is missing.', 'unknown-paragraph');
  const nodes = originNodes(paragraph);
  const sourceOrigin = nodes
    .map((node) => readRunOrigin(runProperties(node), originNamespaceForScope(scope)))
    .find((origin) => origin?.scope === scope);
  if (!sourceOrigin) return part;
  const retained = encodeRetainedPrefixes(sourceOrigin.retainedPrefixes);
  const prior = collectRevisionSites(before);
  const oldIds = new Set(prior.map((site) => site.node.id));
  const oldAddresses = new Set(prior.map((site) => address(site.node)));
  const fresh = new Set(
    collectRevisionSites(part)
      .filter(
        (site) =>
          !oldIds.has(site.node.id) &&
          !oldAddresses.has(address(site.node)) &&
          !site.refused &&
          !site.paragraphMark &&
          (site.node.kind === 'revisionInsert' ||
            site.node.kind === 'revisionDelete' ||
            runPropertySite(part, site.node))
      )
      .map((site) => site.node.id)
  );
  if (!nodes.some((node) => fresh.has(node.id)))
    throw new PackageEditRefusal(
      'A run-origin stage needs new native revision sites.',
      'unsupported'
    );
  // Provenance is not formatting. Never publish it inside a property-history snapshot.
  const snapshots = new Set<string>();
  for (const node of nodes) {
    if (node.kind === 'textValue' || !fresh.has(node.id) || !/PrChange$/.test(node.localName))
      continue;
    for (const child of originNodes(node).slice(1)) snapshots.add(child.id);
  }
  const prefix = originPrefix(
    nodes,
    namespaceBindingsAt(part, paragraph),
    `runOrigin${scope.replace(/-/g, '')}`
  );
  const shells = sourceRunShells(part, nodes, paragraphId, scope);
  const insertedRuns = new Set<string>();
  for (const node of nodes) {
    if (node.kind !== 'revisionInsert' || !fresh.has(node.id)) continue;
    for (const child of originNodes(node)) if (child.kind === 'run') insertedRuns.add(child.id);
  }
  // New opaque controls own their content independently. Their cloned run properties
  // can inherit this preparation's source proof, which must never become control history.
  // Existing controls and every foreign scope stay untouched.
  const bornOpaque = new Set<string>();
  for (const node of nodes) {
    if (!bornOpaque.has(node.id) && (node.kind !== 'contentControl' || findNode(before, node.id)))
      continue;
    bornOpaque.add(node.id);
    for (const child of children(node)) bornOpaque.add(child.id);
  }
  const next = mapOriginNodes(nodes, (node, original) => {
    if (node.kind === 'textValue') return node;
    if (bornOpaque.has(node.id)) {
      let clean = node.attributes.some((attr) => attr.namespaceUri === sourceOrigin.namespaceUri)
        ? withoutOriginMetadata(node, sourceOrigin.namespaceUri)
        : node;
      const oldProperties = runProperties(original);
      const inherited = oldProperties && readRunOrigin(oldProperties, sourceOrigin.namespaceUri);
      if (node.kind === 'run' && inherited?.absent) {
        clean = {
          ...clean,
          children: clean.children.filter(
            (child) =>
              child.id !== oldProperties!.id ||
              child.kind === 'textValue' ||
              child.attributes.length ||
              child.children.length
          ),
        } as OoxmlElement;
      }
      return clean;
    }
    const snapshotOrigin = snapshots.has(node.id) && readRunOrigin(node, sourceOrigin.namespaceUri);
    if (snapshotOrigin && snapshotOrigin.scope === scope)
      return withoutOriginMetadata(node, sourceOrigin.namespaceUri);
    if (node.kind === 'run') {
      const properties = runProperties(node);
      const origin = readRunOrigin(properties, sourceOrigin.namespaceUri);
      if (!properties || !origin || origin.scope !== scope) return node;
      if (node.attributes.some((attr) => attr.namespaceUri === sourceOrigin.namespaceUri))
        throw new PackageEditRefusal('A native run already carries provenance.', 'unsupported');
      const clean = withoutOriginMetadata(properties, origin.namespaceUri);
      const inherited = insertedRuns.has(node.id)
        ? inheritSourceRunShell(node, shells.get(origin.token))
        : node;
      return withOriginMetadata(
        {
          ...inherited,
          children: children(inherited).flatMap((child) =>
            child.id !== properties.id
              ? [child]
              : origin.absent && !clean.children.length && !clean.attributes.length
                ? []
                : [clean]
          ),
        } as OoxmlElement,
        {
          origin: origin.token,
          ...(origin.identity ? { identity: origin.identity } : {}),
          ...(origin.absent ? { absent: '1' } : {}),
          ...(retained ? { retained } : {}),
        },
        prefix
      );
    }
    if (!fresh.has(node.id)) return node;
    if (node.attributes.some((attr) => attr.namespaceUri === sourceOrigin.namespaceUri))
      throw new PackageEditRefusal('A native revision already carries provenance.', 'unsupported');
    return withOriginMetadata(node, { scope, ...(retained ? { retained } : {}) }, prefix);
  });
  if (next === paragraph) return part;
  const made = replaceChildren(part, paragraphId, children(next));
  if (!made.ok) throw new PackageEditRefusal('Run-origin sealing failed.', 'tree-invariant');
  return made.part;
}
