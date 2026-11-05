import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '@docx-editor.dev/core/store';

interface BoundaryRange {
  readonly end: string;
  readonly members: readonly string[];
}

/** Compare exact native range membership. This guard neither repairs nor rewrites references. */
export function blockMovementRangeRefusal(before: OoxmlPart, after: OoxmlPart): string | null {
  const snapshot = (part: OoxmlPart): Map<string, BoundaryRange> | null => {
    const nodes: OoxmlNode[] = [];
    const pending = [{ node: part.root as OoxmlNode, depth: 0 }];
    while (pending.length) {
      const { node, depth } = pending.pop()!;
      if (nodes.length >= 500_000 || depth > 64) return null;
      nodes.push(node);
      if (node.kind !== 'textValue')
        for (let index = node.children.length - 1; index >= 0; index--)
          pending.push({ node: node.children[index]!, depth: depth + 1 });
    }
    const ranges = new Map<string, BoundaryRange>();
    const open = new Map<string, number>();
    const fields: number[] = [];
    let budget = 1_000_000;
    const close = (start: number | undefined, end: number) => {
      if (start === undefined || ranges.size >= 10_000 || (budget -= end - start) < 0) return false;
      ranges.set(nodes[start]!.id, {
        end: nodes[end]!.id,
        members: nodes.slice(start + 1, end).map((node) => node.id),
      });
      return true;
    };
    for (const [index, node] of nodes.entries()) {
      if (node.kind === 'textValue' || node.namespaceUri !== WML_NAMESPACE_URI) continue;
      const value = (name: string) =>
        node.attributes.find(
          (attribute) =>
            attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === name
        )?.value;
      if (node.localName === 'fldChar') {
        const type = value('fldCharType');
        if (type === 'begin') fields.push(index);
        else if (type === 'end' && !close(fields.pop(), index)) return null;
        else if (type === 'separate' && !fields.length) return null;
        continue;
      }
      const boundary =
        /^(bookmark|commentRange|perm|moveFromRange|moveToRange|customXml(?:Ins|Del|MoveFrom|MoveTo)Range)(Start|End)$/.exec(
          node.localName
        );
      if (!boundary) continue;
      const id = value('id');
      if (id === undefined) return null;
      const key = `${boundary[1]}:${/^-?\d{1,10}$/.test(id) ? String(Number(id)) : id}`;
      if (boundary[2] === 'Start') {
        if (open.has(key)) return null;
        open.set(key, index);
      } else {
        if (!close(open.get(key), index)) return null;
        open.delete(key);
      }
    }
    return fields.length || open.size ? null : ranges;
  };
  const original = snapshot(before),
    moved = snapshot(after);
  if (!original || !moved)
    return 'The document contains incomplete or oversized native reference ranges.';
  if (original.size !== moved.size)
    return 'Moving these blocks would change a native field, bookmark, or comment range.';
  for (const [start, range] of original) {
    const next = moved.get(start);
    if (
      !next ||
      next.end !== range.end ||
      next.members.length !== range.members.length ||
      next.members.some((id, index) => id !== range.members[index])
    )
      return 'Moving these blocks would change a native field, bookmark, or comment range.';
  }
  return null;
}
