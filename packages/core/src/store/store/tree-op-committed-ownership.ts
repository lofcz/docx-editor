import { isFldSimple } from '../package/field-nodes.ts';
import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';
import { isInlineContainerProperty } from '../package/inline-container-properties.ts';
import {
  isInlineRunContainer,
  isContentRevisionKind,
  MAX_INLINE_CONTAINER_DEPTH,
  nextInlineContainerDepth,
} from '../package/ooxml-shared.ts';

/** Push pending insertion ownership through one stable container instead of copying that container. */
export function liftedCommittedInsertions(
  nodes: readonly OoxmlNode[],
  inserted: ReadonlySet<string>,
  mint: () => string
): OoxmlNode[] | null {
  const contains = (node: OoxmlNode): boolean =>
    inserted.has(node.id) || (node.kind !== 'textValue' && node.children.some(contains));
  let refused = false;
  const visit = (children: readonly OoxmlNode[], depth: number): OoxmlNode[] =>
    children.flatMap((node): OoxmlNode[] => {
      if (inserted.has(node.id) || !contains(node) || node.kind === 'textValue') return [node];
      if (depth >= MAX_INLINE_CONTAINER_DEPTH) {
        refused = true;
        return [node];
      }
      const inner = visit(node.children, nextInlineContainerDepth(node, depth));
      if (node.kind !== 'revisionInsert') return [{ ...node, children: inner } as OoxmlNode];
      let used = false;
      const distribute = (
        list: readonly OoxmlNode[],
        owner: OoxmlElement,
        nesting: number
      ): OoxmlNode[] => {
        if (nesting >= MAX_INLINE_CONTAINER_DEPTH) {
          refused = true;
          return [...list];
        }
        const output: OoxmlNode[] = [];
        let pending: OoxmlNode[] = [];
        const flush = () => {
          if (!pending.length) return;
          output.push({ ...node, id: used ? mint() : node.id, children: pending } as OoxmlNode);
          used = true;
          pending = [];
        };
        for (const child of list) {
          if (
            isInlineContainerProperty(owner, child) ||
            child.kind === 'contentControlProperties' ||
            child.kind === 'contentControlEndProperties'
          ) {
            flush();
            output.push(child);
          } else if (inserted.has(child.id)) {
            flush();
            output.push(child);
          } else if (contains(child)) {
            flush();
            if (
              child.kind === 'textValue' ||
              !(
                (isInlineRunContainer(child) && !isContentRevisionKind(child.kind)) ||
                isFldSimple(child) ||
                child.kind === 'contentControl' ||
                child.kind === 'contentControlContent'
              )
            ) {
              refused = true;
              output.push(child);
            } else
              output.push({
                ...child,
                children: distribute(
                  child.children,
                  child,
                  nextInlineContainerDepth(child, nesting)
                ),
              } as OoxmlNode);
          } else pending.push(child);
        }
        flush();
        return output;
      };
      return distribute(inner, node, depth);
    });
  const result = visit(nodes, 0);
  return refused ? null : result;
}
