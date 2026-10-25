import type { OoxmlNode } from '../package/ooxml-tree.ts';
import {
  isCanonicalPrimitiveCaptureActive,
  recordNativeTextSlices,
} from '../package/canonical-primitive-capture.ts';
import { cloneWithNewIds } from './tree-op-nodes.ts';

type Leaf = Extract<OoxmlNode, { kind: 'textValue' }>;
function textLeaves(nodes: readonly OoxmlNode[]): Leaf[] {
  const leaves: Leaf[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'text' || node.kind === 'deletedText') {
      for (const child of node.children) if (child.kind === 'textValue') leaves.push(child);
      return;
    }
    if (node.kind !== 'textValue' && node.kind !== 'runProperties')
      for (const child of node.children) visit(child);
  };
  for (const node of nodes) visit(node);
  return leaves;
}

/** Copy native shells while recording exact source positions for every copied Word text leaf. */
export function copyTrackedText(
  nodes: readonly OoxmlNode[],
  sources: readonly OoxmlNode[],
  mint: () => string,
  captureSources = true
): OoxmlNode[] {
  const copies = nodes.map((node) => cloneWithNewIds(node, mint));
  if (!captureSources || !isCanonicalPrimitiveCaptureActive()) return copies;
  const input = textLeaves(sources);
  let index = 0,
    offset = 0;
  for (const target of textLeaves(copies)) {
    const slices: { source: Leaf; start: number; length: number }[] = [];
    let remaining = target.value.length;
    while (remaining) {
      const source = input[index];
      if (!source) throw Error('Tracked text copy exceeds its actual native source.');
      const length = Math.min(remaining, source.value.length - offset);
      if (length) slices.push({ source, start: offset, length });
      remaining -= length;
      offset += length;
      if (offset === source.value.length) {
        index++;
        offset = 0;
      }
    }
    if (slices.length) recordNativeTextSlices(slices, target);
  }
  while (input[index]?.value.length === 0) index++;
  if (index !== input.length || offset)
    throw Error('Tracked text copy omits actual native source text.');
  return copies;
}
