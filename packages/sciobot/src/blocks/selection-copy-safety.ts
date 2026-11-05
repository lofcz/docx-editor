import { WML_NAMESPACE_URI, type OoxmlNode } from '@docx-editor.dev/core/store';

// These children carry document-wide identities, review ownership, or section boundaries.
// A same-story range copy cannot duplicate them by changing only node and paragraph IDs.
const SPECIAL_OWNERS = new Set([
  'drawing',
  'pict',
  'object',
  'altChunk',
  'subDoc',
  'sectPr',
  'dataBinding',
  'footnoteReference',
  'endnoteReference',
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'commentReference',
  'permStart',
  'permEnd',
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
]);

/** Refuse unsupported copy ownership before allocating payloads or changing the package. */
export function blockRangeCopyRefusal(roots: readonly OoxmlNode[]): string | null {
  const pending = roots.map((node) => ({ node, depth: 0 }));
  let count = 0;
  while (pending.length) {
    const { node, depth } = pending.pop()!;
    if (++count > 500_000 || depth > 64) return 'The selected blocks exceed native copy limits.';
    if (node.kind === 'textValue') continue;
    if (
      node.namespaceUri === WML_NAMESPACE_URI &&
      (SPECIAL_OWNERS.has(node.localName) ||
        node.localName.endsWith('Change') ||
        /^customXml(?:Ins|Del|MoveFrom|MoveTo)Range(?:Start|End)$/.test(node.localName))
    )
      return 'The selected blocks contain references or pending changes that require a native copy transaction.';
    for (const child of node.children) pending.push({ node: child, depth: depth + 1 });
  }
  return null;
}
