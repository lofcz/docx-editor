import type { OoxmlNode } from '../package/ooxml-tree.ts';
import { children, readRunOrigin, runProperties } from './run-origin-metadata.ts';
import type { ParagraphOffsetIndex } from './tree-op-segments.ts';

/** Use the actual caret's text element, including when one run has different text shells. */
export function inheritOriginPayload(
  nodes: readonly OoxmlNode[],
  source: OoxmlNode | null | undefined,
  local: number,
  offsets: ParagraphOffsetIndex
): readonly OoxmlNode[] {
  if (!source || !readRunOrigin(runProperties(source))) return nodes;
  let seen = 0;
  let text: OoxmlNode | undefined;
  for (const child of children(source)) {
    if (child.kind === 'runProperties') continue;
    const end = seen + offsets.lengthOf(child);
    if ((child.kind === 'text' || child.kind === 'deletedText') && local >= seen && local <= end) {
      text = child;
      if (local < end) break;
    }
    seen = end;
  }
  if (!text || text.kind === 'textValue') return nodes;
  const shell = text;
  return nodes.map((node) =>
    node.kind === 'text'
      ? ({
          ...node,
          prefix: shell.prefix,
          namespaceBindings: shell.namespaceBindings,
          attributes: shell.attributes,
        } as OoxmlNode)
      : node
  );
}

/** A new tracked run needs its source namespace context before native validation. */
export function inheritOriginRun(node: OoxmlNode, source: OoxmlNode | undefined | null): OoxmlNode {
  if (node.kind !== 'run' || source?.kind !== 'run' || !readRunOrigin(runProperties(source)))
    return node;
  return {
    ...node,
    prefix: source.prefix,
    attributes: source.attributes,
    namespaceBindings: source.namespaceBindings,
  } as OoxmlNode;
}
