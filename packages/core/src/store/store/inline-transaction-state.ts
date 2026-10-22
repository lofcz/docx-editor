import { findNode, parentNodeOf, type TransactionRevisionIds } from '../package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import type { TreeDocOp } from './tree-op-types.ts';
/** Exact native content-revision sites created in one part/transaction, never persisted. */
const states = new WeakMap<TransactionRevisionIds['wrote'], Set<string>>();
export const ownsInlineRevision = (node: OoxmlNode, ids: TransactionRevisionIds | undefined) =>
  !!ids &&
  !!states.get(ids.wrote)?.has(node.id) &&
  (node.kind === 'revisionInsert' || node.kind === 'revisionDelete');
function paragraphId(part: OoxmlPart, op: TreeDocOp): string | null {
  if ('paragraphId' in op) return op.paragraphId;
  if (op.op !== 'setHyperlinkTarget' && op.op !== 'removeHyperlink') return null;
  let node = findNode(part, op.linkId);
  for (let depth = 0; node && depth < 256; depth++) {
    if (node.kind === 'paragraph') return node.id;
    node = parentNodeOf(part, node.id);
  }
  return null;
}
function sites(root: OoxmlNode | null): Set<string> {
  const result = new Set<string>(),
    pending = root ? [root] : [];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'revisionInsert' || node.kind === 'revisionDelete') result.add(node.id);
    if (node.kind !== 'textValue')
      for (const child of node.children) if (child.kind !== 'paragraph') pending.push(child);
  }
  return result;
}
interface InlineSites {
  readonly paragraphId: string;
  readonly sites: ReadonlySet<string>;
}
/** Capture before the edit moves the warm node index to its resulting root. */
export function captureInlineSites(
  part: OoxmlPart,
  op: TreeDocOp,
  ids: TransactionRevisionIds | undefined
): InlineSites | null {
  if (!ids || !('revision' in op) || !op.revision || op.op === 'insertDrawing') return null;
  const id = paragraphId(part, op);
  return id ? { paragraphId: id, sites: sites(findNode(part, id)) } : null;
}
export function rememberInlineSites(
  after: OoxmlPart,
  before: InlineSites | null,
  ids: TransactionRevisionIds | undefined
): void {
  if (!before || !ids) return;
  const next = sites(findNode(after, before.paragraphId));
  let owned = states.get(ids.wrote);
  if (!owned) {
    owned = new Set();
    states.set(ids.wrote, owned);
  }
  for (const site of next) if (!before.sites.has(site)) owned.add(site);
}
