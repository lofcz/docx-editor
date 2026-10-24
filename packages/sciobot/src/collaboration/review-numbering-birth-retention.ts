import * as Y from 'yjs';
import type { ReviewFieldIntent, ReviewFieldTarget } from './review-field-codec';
import { reviewArray } from './review-sequences';

function nativeNodes(doc: Y.Doc): Y.Map<unknown> | null {
  if (!doc.share.has('docx-package-nodes-v1') || !doc.share.has('docx-package-meta-v1'))
    return null;
  const meta = doc.getMap('docx-package-meta-v1');
  if (
    meta.get('protocolVersion') !== 1 ||
    meta.get('sharedSchemaVersion') !== 9 ||
    meta.get('repairVersion') !== 14 ||
    meta.get('canonicalModelVersion') !== 1
  )
    return null;
  return doc.getMap('docx-package-nodes-v1');
}
function wordName(record: unknown): string | null {
  if (!(record instanceof Y.Map) || record.get('deleted') === true) return null;
  const packed = record.get('s') ?? record.get('b');
  if (typeof packed !== 'string') return null;
  const fields = packed.split('\u001f');
  return fields.length === 4 && fields[1] === 'w' ? fields[2]! : null;
}
/** Initial child edges need ownership even when the numbering holder was just allocated. */
export function nativeNumberingChildArray(doc: Y.Doc, array: Y.Array<unknown>): boolean {
  const nodes = nativeNodes(doc),
    field = array._item,
    parent = field?.parent;
  return (
    nodes !== null &&
    field?.parentSub === 'children' &&
    parent instanceof Y.Map &&
    parent._item?.parent === nodes &&
    ['pPr', 'numPr'].includes(wordName(parent) ?? '')
  );
}

/** Retain numbering and paragraph ancestry only while a surviving edit needs it. */
export function retainedNumberingBirths(
  doc: Y.Doc,
  targets: ReadonlyMap<string, ReviewFieldTarget>,
  winners: ReadonlyMap<string, ReviewFieldIntent>
): ReadonlySet<string> {
  const retained = new Set<string>(),
    nodes = nativeNodes(doc);
  if (!nodes) return retained;
  const name = (id: string) => wordName(nodes.get(id));
  // Normalization can remove a cancelled constructor's physical child entries.
  // Its immutable birth targets still prove the ancestry needed by a later peer redo.
  const historical = new Map<string, Set<string>>();
  for (const target of targets.values()) {
    if (!target.sequence || !target.baseline) continue;
    const parent = reviewArray(doc, target)?._item?.parent;
    if (
      !(parent instanceof Y.Map) ||
      parent._item?.parent !== nodes ||
      typeof parent._item.parentSub !== 'string'
    )
      continue;
    const id = parent._item.parentSub,
      edges = historical.get(id) ?? new Set<string>();
    edges.add(target.baseline);
    historical.set(id, edges);
  }
  const children = (id: string): string[] => {
    const record = nodes.get(id);
    const array = record instanceof Y.Map ? record.get('children') : null;
    const live =
      array instanceof Y.Array
        ? array.toArray().filter((value): value is string => typeof value === 'string')
        : [];
    return [...new Set([...live, ...(historical.get(id) ?? [])])];
  };
  const active = new Set<string>();
  for (const [key, winner] of winners) {
    const target = targets.get(key)!;
    if (target.root === 'docx-package-attributes-v1' || target.root === 'docx-package-bindings-v1')
      active.add(target.key.split('\u001f')[0]!);
    else if (target.sequence && winner.value === 'present' && target.baseline) {
      const array = reviewArray(doc, target),
        parent = array?._item?.parent;
      if (
        parent instanceof Y.Map &&
        parent._item?.parent === nodes &&
        typeof parent._item.parentSub === 'string'
      )
        active.add(target.baseline);
    }
  }
  const needed = new Map<string, boolean>();
  const needs = (root: string): boolean => {
    if (needed.has(root)) return needed.get(root)!;
    const visited = new Set<string>(),
      pending = [root];
    while (pending.length) {
      const id = pending.pop()!;
      if (visited.has(id)) continue;
      visited.add(id);
      const record = nodes.get(id);
      if (!(record instanceof Y.Map) || record.get('deleted') === true) continue;
      if (active.has(id)) {
        needed.set(root, true);
        return true;
      }
      pending.push(...children(id));
    }
    needed.set(root, false);
    return false;
  };
  for (const [key, target] of targets) {
    if (!target.sequence || target.initiallyPresent || winners.has(key) || !target.baseline)
      continue;
    const array = reviewArray(doc, target),
      parent = array?._item?.parent;
    if (!(parent instanceof Y.Map) || parent._item?.parent !== nodes) continue;
    const kind = name(target.baseline);
    if ((kind === 'pPr' || kind === 'numPr') && needs(target.baseline)) retained.add(key);
    // A surviving membership still needs its complete original level/reference pair.
    if (
      ['ilvl', 'numId'].includes(kind ?? '') &&
      wordName(parent) === 'numPr' &&
      typeof parent._item.parentSub === 'string' &&
      needs(parent._item.parentSub)
    )
      retained.add(key);
  }
  return retained;
}
