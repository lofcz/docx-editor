import {
  canonicalOoxmlFingerprint,
  createNodeIdAllocator,
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
const PREFIX = 'SciobotTaskRun__';
export interface TaskExportStyle {
  readonly base: string | undefined;
  readonly strikeToggle: boolean;
  readonly color: boolean;
}
const attribute = (node: OoxmlNode, key: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((a) => a.namespaceUri === W && a.localName === key)?.value;
const encode = (value: string) =>
  encodeURIComponent(value).replace(
    /[!'()*_~]/g,
    (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`
  );
export const taskExportStyleKey = (source: TaskExportStyle) =>
  `${source.base === undefined ? 'A' : `S${encode(source.base)}`}_${Number(source.strikeToggle)}${Number(source.color)}`;
function identity(source: TaskExportStyle, suffix = 0) {
  return PREFIX + taskExportStyleKey(source) + (suffix ? `_${suffix}` : '');
}
function sourceOf(id: string): TaskExportStyle | null {
  if (id.length > 512 || !id.startsWith(PREFIX)) return null;
  const match = /^(A|S.*?)_([01])([01])(?:_(\d{1,4}))?$/.exec(id.slice(PREFIX.length));
  if (!match || (match[2] === '0' && match[3] === '0')) return null;
  try {
    const source = {
      base: match[1] === 'A' ? undefined : decodeURIComponent(match[1]!.slice(1)),
      strikeToggle: match[2] === '1',
      color: match[3] === '1',
    };
    return identity(source, Number(match[4] ?? 0)) === id ? source : null;
  } catch {
    return null;
  }
}
function definition(id: string, source: TaskExportStyle) {
  return (
    `<w:style w:type="character" w:customStyle="1" w:styleId="${escapeXml(id)}">` +
    '<w:name w:val="Sciobot Completed Task Run"/><w:aliases w:val="sciobot:task-appearance:v1"/>' +
    (source.base === undefined ? '' : `<w:basedOn w:val="${escapeXml(source.base)}"/>`) +
    '<w:semiHidden/><w:rPr>' +
    (source.strikeToggle ? '<w:strike/>' : '') +
    (source.color ? '<w:color w:val="606060" w:themeColor="text1" w:themeTint="60"/>' : '') +
    '</w:rPr></w:style>'
  );
}
function parsed(name: string, xml: string) {
  const result = readOoxmlPart(`<w:styles xmlns:w="${W}">${xml}</w:styles>`, {
    name,
    contentType: CONTENT_TYPE,
  });
  if (!result.ok) throw new Error('Task export style exceeds native XML limits.');
  return result.part;
}
interface StyleIndex {
  readonly byId: Map<string, OoxmlElement[]>;
  readonly owned: Map<string, TaskExportStyle | null>;
}
const indexes = new WeakMap<object, StyleIndex>();
function indexOf(pkg: OoxmlPackage): StyleIndex {
  const root = stylesPartOf(pkg)?.root;
  if (root && indexes.has(root)) return indexes.get(root)!;
  const index: StyleIndex = { byId: new Map(), owned: new Map() };
  for (const node of root?.children ?? []) {
    if (node.kind === 'textValue' || node.namespaceUri !== W || node.localName !== 'style')
      continue;
    const id = attribute(node, 'styleId');
    if (id === undefined) continue;
    const entries = index.byId.get(id) ?? [];
    entries.push(node);
    index.byId.set(id, entries);
  }
  if (root) indexes.set(root, index);
  return index;
}
/** Metadata, reversible identity, and the complete definition must all match. */
export function taskExportStyleSource(
  pkg: OoxmlPackage,
  id: string | undefined
): TaskExportStyle | null {
  if (!id) return null;
  const source = sourceOf(id);
  if (!source) return null;
  const index = indexOf(pkg);
  if (index.owned.has(id)) return index.owned.get(id)!;
  const matches = index.byId.get(id);
  const expected = parsed('/task-export-style.xml', definition(id, source)).root.children[0]!;
  const result =
    matches?.length === 1 &&
    canonicalOoxmlFingerprint(matches[0]!) === canonicalOoxmlFingerprint(expected)
      ? source
      : null;
  index.owned.set(id, result);
  return result;
}
function fresh(node: OoxmlNode, mint: () => string): OoxmlNode {
  return node.kind === 'textValue'
    ? { ...node, id: mint() }
    : ({
        ...node,
        id: mint(),
        children: node.children.map((child) => fresh(child, mint)),
      } as OoxmlElement);
}
/** Append only definitions required by this detached export. Existing definitions remain exact. */
export function withTaskExportStyles(pkg: OoxmlPackage, sources: readonly TaskExportStyle[]) {
  const index = indexOf(pkg),
    ids = new Map<string, string>(),
    definitions: string[] = [];
  for (const source of sources) {
    const key = taskExportStyleKey(source);
    if (ids.has(key)) continue;
    for (let suffix = 0; suffix < 1024; suffix++) {
      const id = identity(source, suffix);
      if (id.length > 512)
        throw new Error('Task export character style exceeds its identifier limit.');
      if (index.byId.has(id)) {
        if (!taskExportStyleSource(pkg, id)) continue;
      } else definitions.push(definition(id, source));
      ids.set(key, id);
      break;
    }
    if (!ids.has(key)) throw new Error('No task export character style identifier is free.');
  }
  if (!definitions.length) return { pkg, ids };
  const existing = stylesPartOf(pkg),
    directory = pkg.mainDocumentPart.slice(0, pkg.mainDocumentPart.lastIndexOf('/') + 1);
  const occupied = new Set([...pkg.parts.keys(), ...pkg.partBytes.keys()].map(partNameKey));
  let filename = 'sciobot-task-styles.xml';
  for (let index = 1; occupied.has(partNameKey(directory + filename)); index++) {
    if (index >= 1024) throw new Error('No task export styles part name is free.');
    filename = `sciobot-task-styles-${index}.xml`;
  }
  const name = existing?.name ?? directory + filename;
  const seed = parsed(name, definitions.join(''));
  if (existing) {
    const mint = createNodeIdAllocator(existing);
    const inserted = insertChildren(
      existing,
      existing.root.id,
      existing.root.children.length,
      seed.root.children.map((node) => fresh(node, mint))
    );
    if (!inserted.ok) throw new Error('Task export style insertion was refused.');
    return { pkg: withPart(pkg, inserted.part), ids };
  }
  const next = withNewPart(pkg, name, seed.root as OoxmlElement, CONTENT_TYPE);
  const related = withRelationship(next, pkg.mainDocumentPart, RELATIONSHIP, filename);
  if (!related.ok) throw new Error('Task export style relationship was refused.');
  return { pkg: related.pkg, ids };
}
