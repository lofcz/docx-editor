import { MAX_XML_DEPTH, XML_DEFAULT_MAX_ELEMENTS } from '../package/xml-reader.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { PackageEditRefusal } from './package-edit-refusal.ts';
import { normalizedRunOriginParagraph } from './run-origin-resolution.ts';
import { readRevisionOrigins, readRunOrigins } from './run-origin-metadata.ts';

interface SourceEvidence {
  readonly live: readonly string[];
  readonly retained: readonly string[];
  readonly sources: readonly {
    readonly scope: string;
    readonly source: string;
    readonly resolved: boolean;
    readonly prefixes: readonly string[];
  }[];
}
const empty: SourceEvidence = { live: [], retained: [], sources: [] };
const blocks = new WeakMap<OoxmlNode, SourceEvidence>();
const projected = new WeakMap<OoxmlNode, OoxmlNode>();

/** A part-wide liveness check protects scopes split across paragraphs and property histories. */
function evidenceOf(root: OoxmlNode): SourceEvidence {
  let visited = 0;
  const visit = (node: OoxmlNode, depth: number): SourceEvidence => {
    if (++visited > XML_DEFAULT_MAX_ELEMENTS * 2 + 1 || depth > MAX_XML_DEPTH + 1)
      throw new PackageEditRefusal(
        'Run-origin export traversal exceeds its limit.',
        'resource-limit'
      );
    if (node.kind === 'textValue') return empty;
    const memoized = node.kind === 'paragraph' || node.kind === 'tableRow';
    if (memoized) {
      const cached = blocks.get(node);
      if (cached) return cached;
    }
    const live: string[] = [],
      retained: string[] = [],
      sources: SourceEvidence['sources'][number][] = [];
    if (node.kind.startsWith('revision') || /PrChange$/.test(node.localName))
      for (const origin of readRevisionOrigins(node)) {
        live.push(origin.scope);
        for (const prefix of origin.retainedPrefixes) retained.push(prefix);
      }
    if (node.kind === 'run')
      for (const origin of readRunOrigins(node)) {
        if (origin.identity !== 'source') continue;
        sources.push({
          scope: origin.scope,
          source: origin.token.slice(37),
          resolved: origin.resolved,
          prefixes: node.attributes
            .filter(
              (attribute) => attribute.namespaceUri === origin.namespaceUri && attribute.prefix
            )
            .map((attribute) => attribute.prefix!),
        });
      }
    // Formatting snapshots are history data, not current source ownership.
    if (!/PrChange$/.test(node.localName))
      for (const child of node.children) {
        const childEvidence = visit(child, depth + 1);
        for (const scope of childEvidence.live) live.push(scope);
        for (const prefix of childEvidence.retained) retained.push(prefix);
        for (const source of childEvidence.sources) sources.push(source);
      }
    const result =
      live.length || retained.length || sources.length ? { live, retained, sources } : empty;
    if (memoized) blocks.set(node, result);
    return result;
  };
  return visit(root, 0);
}

/**
 * Finish only explicitly resolved proofs whose related scopes have no native revisions left.
 * Every change is a derived immutable node. No canonical primitive or history write runs here.
 */
export function projectCompletedRunOrigins(part: OoxmlPart): OoxmlPart {
  const cached = projected.get(part.root);
  if (cached) return cached === part.root ? part : { ...part, root: cached as OoxmlElement };
  const evidence = evidenceOf(part.root);
  if (!evidence.sources.some((source) => source.resolved)) {
    projected.set(part.root, part.root);
    return part;
  }
  const live = new Set(evidence.live),
    retained = new Set(evidence.retained);
  const bySource = new Map<string, Set<string>>(),
    byScope = new Map<string, Set<string>>();
  for (const source of evidence.sources) {
    if (!source.resolved && !live.has(source.scope)) continue;
    const scopes = bySource.get(source.source) ?? new Set<string>();
    scopes.add(source.scope);
    bySource.set(source.source, scopes);
    const ids = byScope.get(source.scope) ?? new Set<string>();
    ids.add(source.source);
    byScope.set(source.scope, ids);
  }
  const blocked = new Set([
      ...live,
      ...evidence.sources
        .filter((source) => source.prefixes.some((prefix) => retained.has(prefix)))
        .map((source) => source.scope),
    ]),
    pending = [...blocked];
  for (let index = 0; index < pending.length; index++)
    for (const source of byScope.get(pending[index]!) ?? [])
      for (const scope of bySource.get(source) ?? [])
        if (!blocked.has(scope)) {
          blocked.add(scope);
          pending.push(scope);
        }
  const selected = new Set(
    evidence.sources
      .filter((source) => source.resolved && !blocked.has(source.scope))
      .map((source) => source.scope)
  );
  if (!selected.size) {
    projected.set(part.root, part.root);
    return part;
  }
  const visit = (node: OoxmlNode): OoxmlNode => {
    if (node.kind === 'textValue' || /PrChange$/.test(node.localName)) return node;
    const children = node.children.map(visit);
    const current = children.every((child, index) => child === node.children[index])
      ? node
      : ({ ...node, children } as OoxmlElement);
    return current.kind === 'paragraph'
      ? normalizedRunOriginParagraph(current, selected, new Set(), new Set(), [])
      : current;
  };
  const root = visit(part.root) as OoxmlElement;
  projected.set(part.root, root);
  return root === part.root ? part : { ...part, root };
}
