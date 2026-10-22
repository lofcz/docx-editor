import {
  createNodeIdAllocator,
  findNode,
  parentNodeOf,
  replaceChildren,
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
  hasOriginMetadata,
  encodeRetainedPrefixes,
  mapOriginNodes,
  originNodes,
  originPrefix,
  readRunOrigin,
  runProperties,
  withOriginMetadata,
  withoutOriginMetadata,
  RUN_ORIGIN_NS,
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
 * Refuses protected/opaque source interiors and existing provenance on selected runs.
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
  for (const node of nodes)
    if (ids.has(node.id) && originNodes(node).some(hasOriginMetadata))
      throw new PackageEditRefusal(
        'Selected source already carries run-origin metadata.',
        'unsupported'
      );
  const inherited = namespaceBindingsAt(part, paragraph);
  const retained = encodeRetainedPrefixes([
    ...new Set([
      ...[...inherited].filter(([, uri]) => uri === RUN_ORIGIN_NS).map(([prefix]) => prefix),
      ...nodes.flatMap((node) =>
        node.kind === 'textValue'
          ? []
          : node.namespaceBindings
              .filter((binding) => binding.namespaceUri === RUN_ORIGIN_NS)
              .map((binding) => binding.prefix)
      ),
    ]),
  ]);
  const prefix = originPrefix(nodes, inherited);
  const scope = crypto.randomUUID();
  const mint = createNodeIdAllocator(part);
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
    const marked = withOriginMetadata(
      properties,
      {
        origin: `${scope}/${crypto.randomUUID()}`,
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
            { textOrigin: `${scope}/${crypto.randomUUID()}`, space },
            prefix
          );
        }),
      ],
    } as OoxmlNode;
  });
  if (next.kind !== 'paragraph') throw new Error('Invalid run-origin paragraph.');
  const made = replaceChildren(part, paragraphId, next.children);
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
    .map((node) => readRunOrigin(runProperties(node)))
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
  const prefix = originPrefix(nodes, namespaceBindingsAt(part, paragraph));
  const shells = sourceRunShells(part, nodes, paragraphId, scope);
  const insertedRuns = new Set<string>();
  for (const node of nodes) {
    if (node.kind !== 'revisionInsert' || !fresh.has(node.id)) continue;
    for (const child of originNodes(node)) if (child.kind === 'run') insertedRuns.add(child.id);
  }
  const next = mapOriginNodes(nodes, (node) => {
    if (node.kind === 'textValue') return node;
    const snapshotOrigin = snapshots.has(node.id) && readRunOrigin(node);
    if (snapshotOrigin && snapshotOrigin.scope === scope) return withoutOriginMetadata(node);
    if (node.kind === 'run') {
      const properties = runProperties(node);
      const origin = readRunOrigin(properties);
      if (!properties || !origin || origin.scope !== scope) return node;
      if (hasOriginMetadata(node))
        throw new PackageEditRefusal('A native run already carries provenance.', 'unsupported');
      const clean = withoutOriginMetadata(properties);
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
          ...(origin.absent ? { absent: '1' } : {}),
          ...(retained ? { retained } : {}),
        },
        prefix
      );
    }
    if (!fresh.has(node.id)) return node;
    if (hasOriginMetadata(node))
      throw new PackageEditRefusal('A native revision already carries provenance.', 'unsupported');
    return withOriginMetadata(node, { scope, ...(retained ? { retained } : {}) }, prefix);
  });
  if (next === paragraph) return part;
  const made = replaceChildren(part, paragraphId, children(next));
  if (!made.ok) throw new PackageEditRefusal('Run-origin sealing failed.', 'tree-invariant');
  return made.part;
}
