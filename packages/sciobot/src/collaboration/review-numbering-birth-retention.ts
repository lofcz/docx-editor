import * as Y from 'yjs';
import type { ReviewFieldIntent, ReviewFieldTarget } from './review-field-codec';
import { reviewArray } from './review-sequences';

/** Retain numbered-list constructors only while a surviving edit needs their ancestry. */
export function retainedNumberingBirths(
  doc: Y.Doc,
  targets: ReadonlyMap<string, ReviewFieldTarget>,
  winners: ReadonlyMap<string, ReviewFieldIntent>
): ReadonlySet<string> {
  const retained = new Set<string>();
  if (!doc.share.has('docx-package-nodes-v1') || !doc.share.has('docx-package-meta-v1'))
    return retained;
  const meta = doc.getMap('docx-package-meta-v1');
  if (
    meta.get('protocolVersion') !== 1 ||
    meta.get('sharedSchemaVersion') !== 9 ||
    meta.get('repairVersion') !== 14 ||
    meta.get('canonicalModelVersion') !== 1
  )
    return retained;
  const nodes = doc.getMap('docx-package-nodes-v1');
  const name = (id: string): string | null => {
    const record = nodes.get(id);
    if (!(record instanceof Y.Map) || record.get('deleted') === true) return null;
    const packed = record.get('s') ?? record.get('b');
    if (typeof packed !== 'string') return null;
    const fields = packed.split('\u001f');
    return fields.length === 4 && fields[1] === 'w' ? fields[2]! : null;
  };
  const children = (id: string): string[] => {
    const record = nodes.get(id);
    const array = record instanceof Y.Map ? record.get('children') : null;
    return array instanceof Y.Array
      ? array.toArray().filter((value): value is string => typeof value === 'string')
      : [];
  };
  const active = new Set<string>();
  for (const [key] of winners) {
    const target = targets.get(key)!;
    if (target.root === 'docx-package-attributes-v1' || target.root === 'docx-package-bindings-v1')
      active.add(target.key.split('\u001f')[0]!);
    else if (target.sequence) {
      const array = reviewArray(doc, target),
        parent = array?._item?.parent;
      if (
        parent instanceof Y.Map &&
        parent._item?.parent === nodes &&
        typeof parent._item.parentSub === 'string'
      )
        active.add(parent._item.parentSub);
    }
  }
  for (const [key, target] of targets) {
    if (!target.sequence || target.initiallyPresent || winners.has(key) || !target.baseline)
      continue;
    const array = reviewArray(doc, target),
      parent = array?._item?.parent;
    if (!(parent instanceof Y.Map) || parent._item?.parent !== nodes) continue;
    const kind = name(target.baseline);
    if (kind !== 'numPr' && kind !== 'pPr') continue;
    const roots =
      kind === 'numPr'
        ? [target.baseline]
        : children(target.baseline).filter((id) => name(id) === 'numPr');
    const visited = new Set<string>(),
      pending = [...roots];
    while (pending.length) {
      const id = pending.pop()!;
      if (visited.has(id)) continue;
      visited.add(id);
      if (active.has(id)) {
        retained.add(key);
        break;
      }
      pending.push(...children(id));
    }
  }
  return retained;
}
