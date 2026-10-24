import {
  findNode,
  paragraphOffsetIndex,
  parentNodeOf,
  prepareTrackedRunOrigins,
  withPart,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { TrackedStageWriter } from './tracked-stage-writer';

/** Only ordinary source runs need restoration; fresh text and native controls retain their own ownership. */
export function stageSourceRunOrigins(
  writer: TrackedStageWriter,
  partName: string,
  paragraphId: string,
  slices: readonly { readonly start: number; readonly end: number }[]
): string | null {
  const part = writer.currentPackage().parts.get(partName),
    paragraph = part && findNode(part, paragraphId);
  if (!part || paragraph?.kind !== 'paragraph')
    refuseWrite('The source formatting paragraph is missing.', 'invalid-document');
  const ranges: { start: number; end: number }[] = [];
  for (const segment of paragraphOffsetIndex(paragraph).segments) {
    if (segment.removeNodeIds || segment.formatRunIds) continue;
    let node = findNode(part, segment.runId);
    if (node?.kind !== 'run') continue;
    for (let depth = 0; node && node.id !== paragraphId && depth <= 256; depth++) {
      if (node.kind !== 'run' && node.kind !== 'hyperlink') {
        node = null;
        break;
      }
      node = parentNodeOf(part, node.id);
    }
    if (node?.id !== paragraphId) continue;
    for (const slice of slices) {
      const start = Math.max(slice.start, segment.start),
        end = Math.min(slice.end, segment.end);
      if (start < end) ranges.push({ start, end });
    }
  }
  if (!ranges.length) return null;
  let scope: string | null = null;
  writer.edit((pkg) => {
    const prepared = prepareTrackedRunOrigins(pkg.parts.get(partName)!, paragraphId, ranges);
    scope = prepared.scope;
    return withPart(pkg, prepared.part);
  });
  return scope;
}
