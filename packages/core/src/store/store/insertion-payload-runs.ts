import { build } from './tree-op-tracked-builders.ts';
import { inheritOriginPayload, inheritOriginRun } from './run-origin-payload.ts';
import type { OoxmlNode } from '../package/ooxml-tree.ts';
import type { ParagraphOffsetIndex } from './tree-op-segments.ts';

/** A run-child payload inherits the caret formatting; complete runs retain their own formatting. */
export interface InsertionPayload {
  readonly length: number;
  readonly nodes: ((mint: () => string) => readonly OoxmlNode[]) | null;
  readonly runs?: (mint: () => string) => readonly OoxmlNode[];
}

export function insertionPayloadRuns(
  mint: () => string,
  payload: InsertionPayload,
  properties: () => readonly OoxmlNode[],
  source: OoxmlNode | undefined,
  local: number,
  offsets: ParagraphOffsetIndex
): readonly OoxmlNode[] {
  if (!payload.nodes) return payload.runs!(mint);
  const children = [
    ...properties(),
    ...inheritOriginPayload(payload.nodes(mint), source, local, offsets),
  ];
  return [inheritOriginRun(build(mint(), 'run', 'r', [], children), source)];
}
