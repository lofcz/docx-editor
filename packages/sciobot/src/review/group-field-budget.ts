import type { OoxmlNode, OoxmlPackage, OoxmlPart } from '@docx-editor.dev/core/store';
function nodes(part: OoxmlPart) {
  const result = new Map<string, OoxmlNode>(),
    pending: OoxmlNode[] = [part.root];
  while (pending.length) {
    const node = pending.pop()!;
    result.set(node.id, node);
    if (node.kind !== 'textValue') for (const child of node.children) pending.push(child);
  }
  return result;
}
/** Conservative scalar headroom for canonical descriptors, XML attributes and bindings. */
export function partReviewFieldBudget(before: OoxmlPart | undefined, after: OoxmlPart | undefined) {
  if (before === after) return 0;
  const old = before ? nodes(before) : new Map<string, OoxmlNode>(),
    next = after ? nodes(after) : new Map<string, OoxmlNode>();
  let count = 8; // Part-envelope fields do not belong to a canonical child node.
  for (const id of new Set([...old.keys(), ...next.keys()])) {
    const node = next.get(id),
      previous = old.get(id);
    if (previous === node) continue;
    if (!node) {
      if (previous && previous.kind !== 'textValue')
        count += 8 + 2 * (previous.attributes.length + previous.namespaceBindings.length);
      continue;
    }
    if (node.kind === 'textValue') continue;
    // A reordered child can remove an existing sequence entry before its replacement.
    if (
      previous &&
      previous.kind !== 'textValue' &&
      (previous.children.length !== node.children.length ||
        previous.children.some((child, index) => child.id !== node.children[index]?.id))
    )
      count += previous.children.length + node.children.length;
    if (
      !previous ||
      previous.kind === 'textValue' ||
      previous.kind !== node.kind ||
      previous.namespaceUri !== node.namespaceUri ||
      previous.localName !== node.localName ||
      previous.prefix !== node.prefix
    )
      count += 8;
    const attributes = new Map(
      previous && previous.kind !== 'textValue'
        ? previous.attributes.map((a) => [a.namespaceUri + '\0' + a.localName, a])
        : []
    );
    for (const attribute of node.attributes) {
      const key = attribute.namespaceUri + '\0' + attribute.localName,
        value = attributes.get(key);
      if (!value || value.value !== attribute.value || value.prefix !== attribute.prefix)
        count += 2;
      attributes.delete(key);
    }
    count += 2 * attributes.size;
    const bindings = new Map(
      previous && previous.kind !== 'textValue'
        ? previous.namespaceBindings.map((binding) => [binding.prefix, binding.namespaceUri])
        : []
    );
    for (const binding of node.namespaceBindings) {
      if (bindings.get(binding.prefix) !== binding.namespaceUri) count += 2;
      bindings.delete(binding.prefix);
    }
    count += 2 * bindings.size;
  }
  return count;
}
/** Sum stages separately: even a transient payload rewrite can publish a native field. */
export function groupReviewFieldBudget(before: OoxmlPackage, after: OoxmlPackage) {
  let count = 0;
  for (const name of new Set([...before.parts.keys(), ...after.parts.keys()]))
    count += partReviewFieldBudget(before.parts.get(name), after.parts.get(name));
  if (before.relationships !== after.relationships) {
    for (const records of before.relationships.values()) count += 8 * records.length;
    for (const records of after.relationships.values()) count += 8 * records.length;
  }
  if (before.partBytes !== after.partBytes)
    for (const name of new Set([...before.partBytes.keys(), ...after.partBytes.keys()]))
      if (before.partBytes.get(name) !== after.partBytes.get(name)) count += 8;
  if (before.contentTypes !== after.contentTypes)
    count +=
      8 *
      (before.contentTypes.defaults.size +
        before.contentTypes.overrides.size +
        after.contentTypes.defaults.size +
        after.contentTypes.overrides.size);
  return count;
}
