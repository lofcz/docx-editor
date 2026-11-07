import {
  bodyStoryRoot,
  contentControlContentNodeOf,
  findNode,
  parentNodeOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type {
  CollapsedContentControl,
  ContentControlCollapseResult,
} from '../contracts/content-control-collapse.ts';

const MAX_CONTROLS = 4096;
const MAX_NODES = 500_000;
const MAX_DEPTH = 64;

/** Validated ownership and visibility for one local disclosure. @internal */
export interface PreparedContentControlCollapse extends CollapsedContentControl {
  readonly hiddenIds: ReadonlySet<string>;
}

/** Recheck placement even when a moved control keeps its unchanged subtree identity. */
export function collapseOwnerRefusal(
  part: OoxmlPart,
  owner: OoxmlNode
): Extract<ContentControlCollapseResult, { readonly ok: false }> | null {
  const body = bodyStoryRoot(part);
  if (!body) return { ok: false, reason: 'invalid-owner' };
  let ancestor: OoxmlNode | null = owner;
  let depth = 0;
  while (ancestor && ancestor !== body) {
    if (++depth > MAX_DEPTH) return { ok: false, reason: 'limit-exceeded' };
    if (
      !['contentControl', 'contentControlContent', 'table', 'tableRow', 'tableCell'].includes(
        ancestor.kind
      )
    )
      return { ok: false, reason: 'unsupported-content' };
    ancestor = parentNodeOf(part, ancestor.id);
  }
  return ancestor === body ? null : { ok: false, reason: 'invalid-owner' };
}

/** No mutations or callbacks occur while validating an entire replacement view. @internal */
export function prepareContentControlCollapse(
  part: OoxmlPart,
  requests: ReadonlyMap<string, string>,
  budget = { visited: 0 }
):
  | { readonly ok: true; readonly entries: readonly PreparedContentControlCollapse[] }
  | Extract<ContentControlCollapseResult, { readonly ok: false }> {
  if (requests.size > MAX_CONTROLS) return { ok: false, reason: 'limit-exceeded' };
  const body = bodyStoryRoot(part);
  if (!body && requests.size) return { ok: false, reason: 'invalid-owner' };
  const entries: PreparedContentControlCollapse[] = [];
  for (const [controlId, summaryParagraphId] of requests) {
    if (entries.length >= MAX_CONTROLS) return { ok: false, reason: 'limit-exceeded' };
    const owner = typeof controlId === 'string' ? findNode(part, controlId) : null;
    if (owner?.kind !== 'contentControl') return { ok: false, reason: 'invalid-owner' };
    const refusal = collapseOwnerRefusal(part, owner);
    if (refusal) return refusal;
    const content = contentControlContentNodeOf(owner);
    const summary = content?.children[0];
    if (summary?.kind !== 'paragraph' || summary.id !== summaryParagraphId)
      return { ok: false, reason: 'invalid-summary' };
    if (
      content!.children.some(
        (node) => !['paragraph', 'table', 'contentControl'].includes(node.kind)
      )
    )
      return { ok: false, reason: 'unsupported-content' };
    if (content!.children.length > MAX_NODES - budget.visited)
      return { ok: false, reason: 'limit-exceeded' };
    const hiddenIds = new Set<string>();
    const pending = content!.children.map((node) => ({ node, hidden: node !== summary, depth: 0 }));
    while (pending.length) {
      const current = pending.pop()!;
      if (++budget.visited > MAX_NODES || current.depth > MAX_DEPTH)
        return { ok: false, reason: 'limit-exceeded' };
      if (current.hidden) hiddenIds.add(current.node.id);
      if (current.node.kind === 'textValue') continue;
      if (current.node.namespaceUri === WML_NAMESPACE_URI && current.node.localName === 'sectPr')
        return { ok: false, reason: 'unsupported-content' };
      if (pending.length + current.node.children.length > MAX_NODES - budget.visited)
        return { ok: false, reason: 'limit-exceeded' };
      for (const node of current.node.children)
        pending.push({ node, hidden: current.hidden, depth: current.depth + 1 });
    }
    entries.push(Object.freeze({ controlId, summaryParagraphId, hiddenIds }));
  }
  return { ok: true, entries: Object.freeze(entries) };
}
