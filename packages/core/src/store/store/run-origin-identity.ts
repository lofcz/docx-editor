import { sha256FontBytes } from '../package/sha256.ts';
import type { OoxmlNode } from '../package/ooxml-tree.ts';
import { readRunOrigins, readTextOrigins, readRevisionOrigins } from './run-origin-metadata.ts';

const digest = (value: unknown) =>
  sha256FontBytes(new TextEncoder().encode(JSON.stringify(value))).slice(7);

/**
 * Shared replicas have identical native node identities before either proposal.
 * Existing provenance separates later preparation epochs after save/reopen.
 * Never hash document text, formatting, or opaque attribute values.
 * This identity is only a boundary proof; it cannot reconstruct source content.
 */
export function sourceOriginIdentity(nodes: readonly OoxmlNode[]): (node: OoxmlNode) => string {
  const epoch = digest(
    nodes.map((node) => [
      node.id,
      node.kind,
      [
        ...readRunOrigins(node).map((origin) => ['run', origin.namespaceUri, origin.token]),
        ...readTextOrigins(node).map((origin) => ['text', origin.namespaceUri, origin.token]),
        ...readRevisionOrigins(node).map((origin) => [
          'revision',
          origin.namespaceUri,
          origin.scope,
        ]),
      ].sort(),
    ])
  );
  return (node) => {
    const hex = digest([epoch, node.id, node.kind]);
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-8${hex.slice(13, 16)}-a${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
  };
}
