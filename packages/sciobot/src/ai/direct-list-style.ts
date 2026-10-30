import {
  createNodeIdAllocator,
  readOoxmlPart,
  replaceNode,
  withPart,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { LIST_STYLES, type ListStyle } from '../diff/list-types';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const formats: Readonly<Record<ListStyle, string>> = {
  disc: 'bullet',
  circle: 'bullet',
  square: 'bullet',
  decimal: 'decimal',
  'lower-alpha': 'lowerLetter',
  'upper-alpha': 'upperLetter',
  'lower-roman': 'lowerRoman',
  'upper-roman': 'upperRoman',
};
const bullets: Readonly<Record<string, string>> = { disc: '•', circle: '○', square: '■' };
/** Override only the new counter instance. Imported abstract definitions and counters stay intact. */
export function styleListInstance(
  pkg: OoxmlPackage,
  numId: string,
  style: ListStyle
): OoxmlPackage {
  if (!/^\d{1,9}$/.test(numId) || !LIST_STYLES.includes(style))
    throw new Error('Invalid new list style.');
  const part = pkg.parts.get('/word/numbering.xml');
  if (!part) throw new Error('The list numbering part is missing.');
  const num = part.root.children.find(
    (n) =>
      n.kind !== 'textValue' &&
      n.localName === 'num' &&
      n.attributes.some((a) => a.namespaceUri === W && a.localName === 'numId' && a.value === numId)
  );
  if (!num || num.kind === 'textValue') throw new Error('The list counter is missing.');
  // A fresh counter over a reused definition already restarts each level with a start-only
  // override. The style override for a level replaces it and keeps its restart.
  const restarts = new Map<string, string>();
  const startOnly = (child: OoxmlNode) =>
    child.kind !== 'textValue' &&
    child.namespaceUri === W &&
    child.localName === 'lvlOverride' &&
    child.children.length === 1 &&
    child.children[0]!.kind !== 'textValue' &&
    child.children[0]!.localName === 'startOverride';
  for (const child of num.children) {
    if (!startOnly(child) || child.kind === 'textValue') continue;
    const level = child.attributes.find((a) => a.localName === 'ilvl')?.value;
    const start = (child.children[0] as OoxmlElement).attributes.find(
      (a) => a.localName === 'val'
    )?.value;
    if (level !== undefined && start !== undefined && /^\d{1,4}$/.test(start))
      restarts.set(level, start);
  }
  const levels = Array.from(
    { length: 9 },
    (_, level) =>
      `<w:lvlOverride w:ilvl="${level}">${restarts.has(String(level)) ? `<w:startOverride w:val="${restarts.get(String(level))}"/>` : ''}<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="${formats[style]}"/><w:lvlText w:val="${bullets[style] ?? '%' + (level + 1) + '.'}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (level + 1)}" w:hanging="360"/></w:pPr>${bullets[style] ? '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>' : ''}</w:lvl></w:lvlOverride>`
  ).join('');
  const read = readOoxmlPart(
    `<w:numbering xmlns:w="${W}"><w:num w:numId="${numId}">${levels}</w:num></w:numbering>`,
    { name: '/manual-list-style.xml', contentType: part.contentType }
  );
  if (!read.ok) throw new Error('The list style seed is invalid.');
  const seed = read.part.root.children[0]!;
  if (seed.kind === 'textValue') throw new Error('The list seed is missing.');
  const mint = createNodeIdAllocator(part);
  const copy = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'textValue'
      ? { ...node, id: mint() }
      : ({ ...node, id: mint(), children: node.children.map(copy) } as OoxmlNode);
  const result = replaceNode(part, num.id, {
    ...num,
    children: [...num.children.filter((child) => !startOnly(child)), ...seed.children.map(copy)],
  } as OoxmlElement);
  if (!result.ok) throw new Error('The list style override was refused.');
  return withPart(pkg, result.part);
}
