import {
  createNodeIdAllocator,
  insertChildren,
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
/**
 * Give the new counter instance its own styled definition. Imported abstract definitions and
 * counters stay intact.
 *
 * Every `w:num` over one `w:abstractNum` counts on one shared counter, so styling a fresh
 * instance through level overrides on a reused definition would still advance, and be advanced
 * by, the other lists over that definition. A definition of its own keeps the styled list
 * separate, as Word does when a selection takes a different list format.
 */
export function styleListInstance(
  pkg: OoxmlPackage,
  numId: string,
  style: ListStyle
): OoxmlPackage {
  if (!/^\d{1,9}$/.test(numId) || !LIST_STYLES.includes(style))
    throw new Error('Invalid new list style.');
  const part = pkg.parts.get('/word/numbering.xml');
  if (!part) throw new Error('The list numbering part is missing.');
  const wml = (n: OoxmlNode, name: string): n is OoxmlElement =>
    n.kind !== 'textValue' && n.namespaceUri === W && n.localName === name;
  const attribute = (n: OoxmlElement, name: string) =>
    n.attributes.find((a) => a.namespaceUri === W && a.localName === name)?.value;
  const num = part.root.children.find((n) => wml(n, 'num') && attribute(n, 'numId') === numId);
  if (!num || !wml(num, 'num')) throw new Error('The list counter is missing.');
  const abstracts = part.root.children.filter((n) => wml(n, 'abstractNum'));
  // The core allocator stripes w:numId by actor, so reusing that number keeps two peers that
  // style lists at the same time from minting one shared definition ID.
  const used = new Set<number>();
  let next = 0;
  for (const abstract of abstracts) {
    const id = attribute(abstract as OoxmlElement, 'abstractNumId');
    if (id === undefined || !/^\d{1,9}$/.test(id)) continue;
    used.add(Number(id));
    next = Math.max(next, Number(id) + 1);
  }
  const abstractId = used.has(Number(numId)) ? next : Number(numId);
  if (abstractId > 2147483647) throw new Error('No list definition identity is free.');
  const levels = Array.from(
    { length: 9 },
    (_, level) =>
      `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="${formats[style]}"/><w:lvlText w:val="${bullets[style] ?? '%' + (level + 1) + '.'}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (level + 1)}" w:hanging="360"/></w:pPr>${bullets[style] ? '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>' : ''}</w:lvl>`
  ).join('');
  const read = readOoxmlPart(
    `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="${abstractId}"><w:multiLevelType w:val="hybridMultilevel"/>${levels}</w:abstractNum><w:abstractNumId w:val="${abstractId}"/></w:numbering>`,
    { name: '/manual-list-style.xml', contentType: part.contentType }
  );
  if (!read.ok) throw new Error('The list style seed is invalid.');
  const [seed, reference] = read.part.root.children;
  if (!seed || !reference || seed.kind === 'textValue' || reference.kind === 'textValue')
    throw new Error('The list seed is missing.');
  const mint = createNodeIdAllocator(part);
  const copy = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'textValue'
      ? { ...node, id: mint() }
      : ({ ...node, id: mint(), children: node.children.map(copy) } as OoxmlNode);
  // Only the definition reference changes; the fresh instance keeps its restart overrides.
  const current = num.children.find((child) => wml(child, 'abstractNumId'));
  if (!current) throw new Error('The list counter has no definition.');
  let lastAbstract = -1;
  part.root.children.forEach((n, index) => {
    if (wml(n, 'abstractNum')) lastAbstract = index;
  });
  const after =
    lastAbstract >= 0
      ? lastAbstract + 1
      : Math.max(
          0,
          part.root.children.findIndex((n) => wml(n, 'num'))
        );
  const inserted = insertChildren(part, part.root.id, after, [copy(seed)]);
  if (!inserted.ok) throw new Error('The list style definition was refused.');
  const result = replaceNode(inserted.part, current.id, {
    ...reference,
    id: current.id,
  } as OoxmlElement);
  if (!result.ok) throw new Error('The list style definition was refused.');
  return withPart(pkg, result.part);
}
