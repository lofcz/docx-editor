import { WML_NAMESPACE_URI, type OoxmlNode } from '../package/ooxml-tree.ts';
const kids = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
const word = (node: OoxmlNode, name: string) =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;

/**
 * The face a NEW run takes from the run it is being inserted into — never that run's pending
 * FORMAT RECORD.
 *
 * A split copies the record into both halves, correctly: the same characters are still under
 * the same proposal. A run being inserted is different. Its characters were never in the state
 * the record describes, so inheriting it put them under somebody's pending decision — reject
 * the format card and the words just typed came back in a colour nobody had ever given them,
 * while their own insertion card was still unanswered. The mirror of `insideOwnInsertion`,
 * which refuses to WRITE a record in the same position.
 */
export function insertedRunProperties(mint: () => string, run: OoxmlNode): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  for (const child of kids(run)) {
    if (child.kind !== 'runProperties') continue;
    const kept = kids(child).filter((entry) => !word(entry, 'rPrChange'));
    // A container holding nothing but the record is not a face to inherit.
    if (kept.length === 0 && kids(child).length !== 0) continue;
    out.push(copy(mint, { ...child, children: kept } as OoxmlNode));
  }
  return out;
}

/** Deep copy with fresh ids, so a split run's halves are two nodes and not one twice. */
export function copy(mint: () => string, node: OoxmlNode): OoxmlNode {
  if (node.kind === 'textValue') return { id: mint(), kind: 'textValue', value: node.value };
  return {
    ...node,
    id: mint(),
    children: node.children.map((child) => copy(mint, child)),
  } as OoxmlNode;
}
