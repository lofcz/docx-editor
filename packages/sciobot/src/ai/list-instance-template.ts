import {
  createNodeIdAllocator,
  replaceNode,
  insertChildren,
  withPart,
  type OoxmlNode,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const named = (node: OoxmlNode, name: string): node is OoxmlElement =>
  node.kind !== 'textValue' && node.namespaceUri === W && node.localName === name;
const value = (node: OoxmlElement, name: string) =>
  node.attributes.find((attribute) => attribute.namespaceUri === W && attribute.localName === name)
    ?.value;
function counter(pkg: OoxmlPackage, id: string) {
  if (!/^\d{1,9}$/.test(id) || Number(id) <= 0)
    throw new Error('Invalid numbering instance identity.');
  const part = pkg.parts.get('/word/numbering.xml');
  if (!part) throw new Error('The numbering part is missing.');
  const found = part.root.children.filter(
    (node) => named(node, 'num') && value(node, 'numId') === id
  );
  if (found.length !== 1 || !named(found[0]!, 'num'))
    throw new Error('The numbering instance is missing or ambiguous.');
  return { part, node: found[0]! as OoxmlElement };
}
const levelOf = (node: OoxmlNode) =>
  named(node, 'lvlOverride') ? value(node, 'ilvl')?.trim() : undefined;
/**
 * Every `w:num` over one abstract definition counts on one shared counter, so a fresh
 * instance restarts each level with a start-only override. A copied template keeps those
 * restarts for the levels where it does not set its own start.
 */
function withRestarts(
  copied: readonly OoxmlNode[],
  fresh: readonly OoxmlNode[],
  copy: (node: OoxmlNode) => OoxmlNode
): OoxmlNode[] {
  const restarts = new Map<string, OoxmlElement>();
  for (const node of fresh) {
    const level = levelOf(node);
    const start = named(node, 'lvlOverride')
      ? node.children.find((c) => named(c, 'startOverride'))
      : undefined;
    if (level !== undefined && start && named(start, 'startOverride')) restarts.set(level, start);
  }
  const covered = new Set<string>();
  const out = copied.map((node) => {
    const level = levelOf(node);
    if (level === undefined || !named(node, 'lvlOverride')) return node;
    covered.add(level);
    const restart = restarts.get(level);
    if (!restart || node.children.some((child) => named(child, 'startOverride'))) return node;
    return { ...node, children: [copy(restart), ...node.children] } as OoxmlElement;
  });
  for (const node of fresh) {
    const level = levelOf(node);
    if (level !== undefined && restarts.has(level) && !covered.has(level)) out.push(copy(node));
  }
  return out;
}
/** Fork a counter without replacing its imported abstract or level formatting. */
export function copyListInstanceTemplate(
  pkg: OoxmlPackage,
  sourceId: string,
  destinationId: string
): OoxmlPackage {
  const source = counter(pkg, sourceId),
    destination = counter(pkg, destinationId),
    mint = createNodeIdAllocator(destination.part);
  const copy = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'textValue'
      ? { ...node, id: mint() }
      : ({ ...node, id: mint(), children: node.children.map(copy) } as OoxmlElement);
  const replacement = {
    ...source.node,
    id: destination.node.id,
    attributes: source.node.attributes.map((attribute) =>
      attribute.namespaceUri === W && attribute.localName === 'numId'
        ? { ...attribute, value: destinationId }
        : attribute
    ),
    children: withRestarts(source.node.children.map(copy), destination.node.children, copy),
  };
  const result = replaceNode(destination.part, destination.node.id, replacement as OoxmlElement);
  if (!result.ok) throw new Error('The imported counter template cannot be copied.');
  return withPart(pkg, result.part);
}
/** Replace one existing start override, retaining every other level and opaque field. */
export function setListInstanceStart(
  pkg: OoxmlPackage,
  id: string,
  level: number,
  start: number
): OoxmlPackage {
  if (
    !Number.isInteger(level) ||
    level < 0 ||
    level > 8 ||
    !Number.isSafeInteger(start) ||
    start < 0 ||
    start > 2147483647
  )
    throw new Error('Invalid numbering start or level.');
  const { part, node } = counter(pkg, id),
    mint = createNodeIdAllocator(part);
  const attr = (localName: string, n: number) => ({
    kind: 'genericExtension' as const,
    namespaceUri: W,
    prefix: 'w',
    localName,
    value: String(n),
  });
  const make = (
    localName: string,
    attributes: OoxmlElement['attributes'],
    children: readonly OoxmlNode[]
  ): OoxmlElement => ({
    kind: 'generic',
    id: mint(),
    namespaceUri: W,
    prefix: 'w',
    localName,
    namespaceBindings: [{ prefix: 'w', namespaceUri: W }],
    attributes,
    children,
  });
  const overrides = node.children.filter((child) => {
    if (!named(child, 'lvlOverride')) return false;
    const lexical = value(child, 'ilvl')?.trim();
    return lexical !== undefined && /^[+]?\d+$/.test(lexical) && Number(lexical) === level;
  });
  if (overrides.length > 1) throw new Error('The numbering level has ambiguous overrides.');
  const existing = overrides[0];
  if (!existing) {
    const inserted = insertChildren(part, node.id, node.children.length, [
      make('lvlOverride', [attr('ilvl', level)], [make('startOverride', [attr('val', start)], [])]),
    ]);
    if (!inserted.ok) throw new Error('The new numbering start was refused.');
    return withPart(pkg, inserted.part);
  }
  if (existing.kind === 'textValue') throw new Error('Invalid numbering override.');
  const starts = existing.children.filter((child) => named(child, 'startOverride'));
  if (starts.length > 1) throw new Error('The numbering start is ambiguous.');
  const old = starts[0];
  if (
    old &&
    old.kind !== 'textValue' &&
    !old.attributes.some(
      (attribute) => attribute.namespaceUri === W && attribute.localName === 'val'
    )
  )
    throw new Error('The numbering start has no Word value.');
  const children = old
    ? existing.children.map((child) =>
        child.id !== old.id || child.kind === 'textValue'
          ? child
          : ({
              ...child,
              attributes: child.attributes.map((attribute) =>
                attribute.namespaceUri === W && attribute.localName === 'val'
                  ? { ...attribute, value: String(start) }
                  : attribute
              ),
            } as OoxmlElement)
      )
    : [make('startOverride', [attr('val', start)], []), ...existing.children];
  const result = replaceNode(part, existing.id, { ...existing, children } as OoxmlElement);
  if (!result.ok) throw new Error('The numbering start was refused.');
  return withPart(pkg, result.part);
}
