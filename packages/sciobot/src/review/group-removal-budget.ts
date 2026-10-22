import type { OoxmlNode, OoxmlPackage, OoxmlPart } from '@docx-editor.dev/core/store';

function nodeIds(root: OoxmlNode): Set<string> {
  const ids = new Set<string>(),
    pending = [root];
  while (pending.length) {
    const node = pending.pop()!;
    ids.add(node.id);
    if (node.kind !== 'textValue') for (const child of node.children) pending.push(child);
  }
  return ids;
}

/**
 * A retained shared record has at most one removal flag. Count removed canonical nodes,
 * including property snapshots and opaque application XML, before any native publication.
 * Text-only removals may reserve unused capacity. Unchanged parts allocate no ID sets.
 */
export function partRemovalBudget(before: OoxmlPart, after: OoxmlPart | undefined): number {
  if (before === after) return 0;
  const retained = after ? nodeIds(after.root) : new Set<string>();
  let count = 0;
  for (const id of nodeIds(before.root)) if (!retained.has(id)) count++;
  return count;
}

export function groupRemovalBudget(before: OoxmlPackage, after: OoxmlPackage): number {
  let count = 0;
  for (const [name, part] of before.parts) {
    count += partRemovalBudget(part, after.parts.get(name));
  }
  return count;
}
