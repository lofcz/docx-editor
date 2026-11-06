import {
  inspectDrawingSupport,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { nativeImageResource } from './native-image-fields';
import { blockRangeCopyRefusal } from './selection-copy-safety';

const WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';

/** Same-story embedded pictures retain resource relationships and need independent drawing identities. */
export function drawingCopyRefusal(
  pkg: OoxmlPackage,
  partName: string,
  roots: readonly OoxmlNode[]
): string | null {
  const drawings: OoxmlElement[] = [];
  const pending = roots.map((node) => ({ node, depth: 0 }));
  let count = 0;
  while (pending.length) {
    const { node, depth } = pending.pop()!;
    if (++count > 500_000 || depth > 64) return 'The selected blocks exceed native copy limits.';
    if (node.kind === 'textValue') continue;
    if (node.namespaceUri === WML_NAMESPACE_URI) {
      if (['pict', 'object', 'altChunk', 'subDoc'].includes(node.localName))
        return 'Legacy drawings and embedded objects require a separate native copy operation.';
      if (node.localName === 'drawing') {
        if (node.kind !== 'drawing') return 'The copied drawing has no supported native identity.';
        drawings.push(node);
      }
    }
    if (
      node.attributes.some(
        (attribute) =>
          attribute.namespaceUri === WP14 && ['anchorId', 'editId'].includes(attribute.localName)
      )
    )
      return 'The copied drawing uses unsupported placement identities.';
    for (const child of node.children) pending.push({ node: child, depth: depth + 1 });
  }
  if (!drawings.length) return null;
  const references = blockRangeCopyRefusal(roots);
  if (references) return references;
  const report = inspectDrawingSupport(pkg);
  if (!report.complete) return 'The drawing inventory exceeds native copy limits.';
  const entries = new Map(
    report.entries
      .filter((entry) => entry.ownerPartName === partName && entry.nodeId === entry.drawingNodeId)
      .map((entry) => [entry.nodeId, entry])
  );
  for (const drawing of drawings) {
    const entry = entries.get(drawing.id);
    const resource = nativeImageResource(pkg, partName, drawing);
    if (
      drawing.children.length !== 1 ||
      drawing.children[0]?.kind !== 'inlineDrawing' ||
      !entry ||
      entry.graphicKind !== 'picture' ||
      entry.source !== drawing ||
      !resource ||
      (!pkg.partBytes.has(resource.partName) && !pkg.partBytes.has(resource.partName.slice(1)))
    )
      return 'Only embedded inline pictures support this native copy operation.';
  }
  return null;
}
