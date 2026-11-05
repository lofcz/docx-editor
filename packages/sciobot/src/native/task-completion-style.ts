import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  createNodeIdAllocator,
  directParagraphProperties,
  escapeXml,
  insertChildren,
  partNameKey,
  readOoxmlPart,
  stylesPartOf,
  withNewPart,
  withPart,
  withRelationship,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const PREFIX = 'SciobotTaskDone__';
const NAME = 'Sciobot Completed Task';
type Source = { base: string | undefined };
const sourceCache = new WeakMap<object, Map<string, Source | null>>();
interface StyleIndex {
  readonly definitions: readonly OoxmlElement[];
  readonly byId: ReadonlyMap<string, readonly OoxmlElement[]>;
  readonly defaultId: string | undefined;
}
const indexCache = new WeakMap<object, StyleIndex>();
const attribute = (node: OoxmlNode | undefined, key: string) =>
  !node || node.kind === 'textValue'
    ? undefined
    : node.attributes.find((a) => a.namespaceUri === W && a.localName === key)?.value;
const is = (node: OoxmlNode, name: string): node is OoxmlElement =>
  node.kind !== 'textValue' && node.namespaceUri === W && node.localName === name;
function styleIndex(pkg: OoxmlPackage): StyleIndex {
  const root = stylesPartOf(pkg)?.root;
  if (root && indexCache.has(root)) return indexCache.get(root)!;
  const styles: OoxmlElement[] = [];
  const byId = new Map<string, OoxmlElement[]>();
  let defaultId: string | undefined;
  for (const node of root?.children ?? []) {
    if (!is(node, 'style')) continue;
    styles.push(node);
    const id = attribute(node, 'styleId');
    if (id === undefined) continue;
    const entries = byId.get(id) ?? [];
    entries.push(node);
    byId.set(id, entries);
    if (
      defaultId === undefined &&
      attribute(node, 'type') === 'paragraph' &&
      ['1', 'true', 'on'].includes(attribute(node, 'default') ?? '')
    )
      defaultId = id;
  }
  const index = { definitions: styles, byId, defaultId };
  if (root) indexCache.set(root, index);
  return index;
}
const encode = (value: string) =>
  encodeURIComponent(value).replace(
    /[!'()*_~]/g,
    (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`
  );
function identity(base: string | undefined, suffix = 0): string {
  return PREFIX + (base === undefined ? 'A' : 'S' + encode(base)) + (suffix ? `_${suffix}` : '');
}
function source(id: string): Source | null {
  if (!id.startsWith(PREFIX)) return null;
  const match = /^(A|S.*?)(?:_(\d+))?$/.exec(id.slice(PREFIX.length));
  if (!match) return null;
  try {
    const base = match[1] === 'A' ? undefined : decodeURIComponent(match[1]!.slice(1));
    return identity(base, Number(match[2] ?? 0)) === id ? { base } : null;
  } catch {
    return null;
  }
}
function definition(id: string, inherited: string | undefined): string {
  return (
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${escapeXml(id)}">` +
    `<w:name w:val="${NAME}"/>` +
    (inherited === undefined ? '' : `<w:basedOn w:val="${escapeXml(inherited)}"/>`) +
    '<w:semiHidden/><w:rPr><w:strike/><w:color w:val="808080" w:themeColor="text1" w:themeTint="80"/></w:rPr></w:style>'
  );
}
function parsed(name: string, xml: string) {
  const result = readOoxmlPart(`<w:styles xmlns:w="${W}">${xml}</w:styles>`, {
    name,
    contentType: CONTENT_TYPE,
  });
  if (!result.ok) throw new Error('The task completion style is invalid.');
  return result.part;
}
/** An imported style qualifies only when its complete definition matches the owned style. */
export function taskCompletionStyleSource(
  pkg: OoxmlPackage,
  id: string | undefined
): Source | null {
  if (!id) return null;
  const original = source(id);
  if (!original) return null;
  const root = stylesPartOf(pkg)?.root;
  if (!root) return null;
  let cached = sourceCache.get(root);
  if (!cached) {
    cached = new Map();
    sourceCache.set(root, cached);
  }
  if (cached.has(id)) return cached.get(id)!;
  const matches = styleIndex(pkg).byId.get(id);
  if (!matches?.length) return null;
  const result = matches.every((style) => {
    const inherited =
      original.base ??
      attribute(
        style.children.find((n) => is(n, 'basedOn')),
        'val'
      );
    const expected = parsed('/task-style.xml', definition(id, inherited)).root.children[0]!;
    return canonicalOoxmlFingerprint(style) === canonicalOoxmlFingerprint(expected);
  })
    ? original
    : null;
  cached.set(id, result);
  return result;
}
function fresh(node: OoxmlNode, next: () => string): OoxmlNode {
  return node.kind === 'textValue'
    ? { ...node, id: next() }
    : ({
        ...node,
        id: next(),
        children: node.children.map((child) => fresh(child, next)),
      } as OoxmlElement);
}
function ensureStyles(pkg: OoxmlPackage, bases: Iterable<string | undefined>) {
  const index = styleIndex(pkg),
    ids = new Map<string | undefined, string>(),
    xml: string[] = [];
  for (const base of bases) {
    for (let suffix = 0; suffix < 1024; suffix++) {
      const id = identity(base, suffix);
      if (id.length > 512) throw new Error('The task source style identifier exceeds its limit.');
      if (index.byId.has(id)) {
        if (!taskCompletionStyleSource(pkg, id)) continue;
      } else xml.push(definition(id, base ?? index.defaultId));
      ids.set(base, id);
      break;
    }
    if (!ids.has(base)) throw new Error('No task completion style identifier is free.');
  }
  if (!xml.length) return { pkg, ids };
  const existing = stylesPartOf(pkg),
    directory = pkg.mainDocumentPart.slice(0, pkg.mainDocumentPart.lastIndexOf('/') + 1);
  const occupied = new Set([...pkg.parts.keys(), ...pkg.partBytes.keys()].map(partNameKey));
  let filename = 'styles.xml';
  for (let n = 1; occupied.has(partNameKey(directory + filename)); n++)
    filename = `sciobot-task-styles-${n}.xml`;
  const name = existing?.name ?? directory + filename;
  // The native parser applies its XML element and size limits to the complete seed.
  const seed = parsed(name, xml.join(''));
  if (existing) {
    const mint = createNodeIdAllocator(existing);
    const inserted = insertChildren(
      existing,
      existing.root.id,
      existing.root.children.length,
      seed.root.children.map((node) => fresh(node, mint))
    );
    if (!inserted.ok) throw new Error('The task completion style was refused.');
    return { pkg: withPart(pkg, inserted.part), ids };
  }
  const next = withNewPart(pkg, name, seed.root as OoxmlElement, CONTENT_TYPE);
  const related = withRelationship(next, pkg.mainDocumentPart, RELATIONSHIP, filename);
  if (!related.ok) throw new Error('The task completion style relationship was refused.');
  return { pkg: related.pkg, ids };
}
/** Seed before room bootstrap so shared history never owns a style used by another participant. */
export function withTaskCompletionStyles(pkg: OoxmlPackage): OoxmlPackage {
  const bases = new Set<string | undefined>([undefined]);
  for (const style of styleIndex(pkg).definitions) {
    const id = attribute(style, 'styleId');
    if (id && attribute(style, 'type') === 'paragraph' && !taskCompletionStyleSource(pkg, id))
      if (identity(id).length <= 512) bases.add(id);
  }
  return ensureStyles(pkg, bases).pkg;
}
/** Shared edits may reference initial definitions, but cannot create their shared lifetime. */
export function addsTaskCompletionStyles(before: OoxmlPackage, after: OoxmlPackage): boolean {
  if (stylesPartOf(before) === stylesPartOf(after)) return false;
  return styleIndex(after).definitions.some((style) => {
    const id = attribute(style, 'styleId');
    return !!taskCompletionStyleSource(after, id) && !taskCompletionStyleSource(before, id);
  });
}
export const TASK_COMPLETION_PREPARATION_ERROR =
  'Prepare task completion styles before starting collaboration.';
/** Change only the paragraph style reference. Authored run properties remain untouched. */
export function applyTaskCompletionStyle(
  pkg: OoxmlPackage,
  partName: string,
  paragraphId: string,
  checked: boolean
): OoxmlPackage {
  const part = pkg.parts.get(partName)!;
  const properties = directParagraphProperties(part, paragraphId);
  const current = properties.find((p) => p.localName === 'pStyle')?.attributes?.val;
  const original = taskCompletionStyleSource(pkg, current);
  if ((checked && original) || (!checked && !original)) return pkg;
  const ensured = checked ? ensureStyles(pkg, [current]) : null;
  const desired = ensured
    ? { pkg: ensured.pkg, id: ensured.ids.get(current)! }
    : { pkg, id: original!.base };
  const result = applyTreeOp(desired.pkg.parts.get(partName)!, {
    op: 'setParagraphProperties',
    paragraphId,
    properties: [
      ...properties.filter((p) => p.localName !== 'pStyle'),
      ...(desired.id === undefined
        ? []
        : [{ localName: 'pStyle', attributes: { val: desired.id } }]),
    ],
  });
  if (!result.ok) throw new Error('The task completion paragraph style was refused.');
  return withPart(desired.pkg, result.part);
}
