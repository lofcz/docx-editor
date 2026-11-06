import {
  resolveContentType,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';

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
  // Native binary descriptors also carry one owned boolean: retainedOverride. Their
  // publication integrates that flag even when the command removes no canonical node.
  for (const [name, bytes] of after.partBytes) {
    const type = resolveContentType(after.contentTypes, name);
    if (!type.ok || !type.contentType.startsWith('image/')) continue;
    const normalized = '/' + name.replace(/^\//, '');
    if (
      before.partBytes.get(name) !== bytes ||
      before.contentTypes.overrides.get(normalized) !== after.contentTypes.overrides.get(normalized)
    )
      count++;
  }
  return count;
}
