// Restore only exact paragraph styles from the earlier experimental implementation.
// Live editing never creates these definitions or changes paragraph style references.
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  escapeXml,
  readOoxmlPart,
  stylesPartOf,
  withPart,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const PREFIX = 'SciobotTaskDone__';
const NAME = 'Sciobot Completed Task';
type Source = { base: string | undefined };
const sourceCache = new WeakMap<object, Map<string, Source | null>>();
interface StyleIndex {
  readonly byId: ReadonlyMap<string, readonly OoxmlElement[]>;
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
  const byId = new Map<string, OoxmlElement[]>();
  for (const node of root?.children ?? []) {
    if (!is(node, 'style')) continue;
    const id = attribute(node, 'styleId');
    if (id === undefined) continue;
    const entries = byId.get(id) ?? [];
    entries.push(node);
    byId.set(id, entries);
  }
  const index = { byId };
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
  if (id.length > 512 || !id.startsWith(PREFIX)) return null;
  const match = /^(A|S.*?)(?:_(\d+))?$/.exec(id.slice(PREFIX.length));
  if (!match) return null;
  try {
    const base = match[1] === 'A' ? undefined : decodeURIComponent(match[1]!.slice(1));
    return identity(base, Number(match[2] ?? 0)) === id ? { base } : null;
  } catch {
    return null;
  }
}
function definition(id: string, inherited: string | undefined, strike = true): string {
  return (
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${escapeXml(id)}">` +
    `<w:name w:val="${NAME}"/>` +
    (inherited === undefined ? '' : `<w:basedOn w:val="${escapeXml(inherited)}"/>`) +
    '<w:semiHidden/><w:rPr>' +
    (strike ? '<w:strike/>' : '') +
    '<w:color w:val="606060" w:themeColor="text1" w:themeTint="60"/></w:rPr></w:style>'
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
    const runProperties = style.children.find((node) => is(node, 'rPr'));
    const strike =
      runProperties?.kind !== 'textValue' &&
      runProperties?.children.some((node) => is(node, 'strike'));
    const expected = parsed('/task-style.xml', definition(id, inherited, !!strike)).root
      .children[0]!;
    return canonicalOoxmlFingerprint(style) === canonicalOoxmlFingerprint(expected);
  })
    ? original
    : null;
  cached.set(id, result);
  return result;
}
/** Task syntax already carries completion. Detached Markdown reads retain authored formatting only. */
export function withoutTaskCompletionAppearance(pkg: OoxmlPackage): OoxmlPackage {
  const part = stylesPartOf(pkg);
  if (!part) return pkg;
  let changed = false;
  const children = part.root.children.map((node) => {
    if (!is(node, 'style') || !taskCompletionStyleSource(pkg, attribute(node, 'styleId')))
      return node;
    changed = true;
    return {
      ...node,
      children: node.children.filter((child) => !is(child, 'rPr')),
    } as OoxmlElement;
  });
  return changed
    ? withPart(pkg, { ...part, root: { ...part.root, children } as OoxmlElement })
    : pkg;
}
/** Detached preparation restores absent and explicit paragraph references exactly. */
export function restoreLegacyTaskCompletionStyle(
  pkg: OoxmlPackage,
  partName: string,
  paragraphId: string
): OoxmlPackage {
  const part = pkg.parts.get(partName)!;
  const properties = directParagraphProperties(part, paragraphId);
  const current = properties.find((p) => p.localName === 'pStyle')?.attributes?.val;
  const original = taskCompletionStyleSource(pkg, current);
  if (!original) return pkg;
  const result = applyTreeOp(part, {
    op: 'setParagraphProperties',
    paragraphId,
    properties: [
      ...properties.filter((p) => p.localName !== 'pStyle'),
      ...(original.base === undefined
        ? []
        : [{ localName: 'pStyle', attributes: { val: original.base } }]),
    ],
  });
  if (!result.ok) throw new Error('The legacy task paragraph style restoration was refused.');
  return withPart(pkg, result.part);
}
