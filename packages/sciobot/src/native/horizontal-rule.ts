import {
  createNodeIdAllocator,
  insertChildren,
  partNameKey,
  readOoxmlPart,
  resolveRelationship,
  stylesPartOf,
  withNewPart,
  withPart,
  withRelationship,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';

/** Explicit application identity; ordinary imported paragraph borders never imply a rule. */
export const SCIOBOT_HORIZONTAL_RULE_STYLE_ID = 'SciobotHorizontalRule';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const STYLE = `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${SCIOBOT_HORIZONTAL_RULE_STYLE_ID}"><w:name w:val="Sciobot Horizontal Rule"/><w:next w:val="Normal"/></w:style>`;
/** A native Word paragraph border: 0.75 pt single closing edge, including grouped paragraphs, with a one-point text gap. */
export const HORIZONTAL_RULE_PARAGRAPH_XML = `<w:p><w:pPr><w:pStyle w:val="${SCIOBOT_HORIZONTAL_RULE_STYLE_ID}"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/><w:between w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;

const attr = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)
        ?.value;
const is = (node: OoxmlNode, name: string): node is OoxmlElement =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;

function exactAttributes(node: OoxmlElement, expected: Record<string, string>): boolean {
  return (
    node.attributes.length === Object.keys(expected).length &&
    node.attributes.every(
      (attribute) =>
        attribute.namespaceUri === WML_NAMESPACE_URI &&
        expected[attribute.localName] === attribute.value
    )
  );
}
function emptyProperty(node: OoxmlNode, name: string, expected: Record<string, string>): boolean {
  return is(node, name) && node.children.length === 0 && exactAttributes(node, expected);
}

/** Strict definition matching refuses imported collisions without deleting their metadata. */
export function horizontalRuleStyleState(pkg: OoxmlPackage): 'missing' | 'present' | 'conflict' {
  const relations = (pkg.relationships.get(pkg.mainDocumentPart) ?? []).filter(
    (r) => r.type === RELATIONSHIP
  );
  if (relations.length > 1) return 'conflict';
  const part = stylesPartOf(pkg);
  if (relations.length) {
    const resolved = resolveRelationship(relations[0]!);
    if (
      resolved.mode !== 'Internal' ||
      !resolved.target.ok ||
      !part ||
      part.name !== resolved.target.partName
    )
      return 'conflict';
  }
  if (part && !is(part.root, 'styles')) return 'conflict';
  const matches =
    part?.root.children.filter(
      (node) => is(node, 'style') && attr(node, 'styleId') === SCIOBOT_HORIZONTAL_RULE_STYLE_ID
    ) ?? [];
  if (!matches.length) return 'missing';
  if (matches.length !== 1) return 'conflict';
  const style = matches[0]!;
  return is(style, 'style') &&
    exactAttributes(style, {
      type: 'paragraph',
      customStyle: '1',
      styleId: SCIOBOT_HORIZONTAL_RULE_STYLE_ID,
    }) &&
    style.children.length === 2 &&
    style.children.some((node) =>
      emptyProperty(node, 'name', { val: 'Sciobot Horizontal Rule' })
    ) &&
    style.children.some((node) => emptyProperty(node, 'next', { val: 'Normal' }))
    ? 'present'
    : 'conflict';
}

/** Only the authored empty marker qualifies; unknown properties/content fail closed. */
export function isHorizontalRuleParagraph(node: OoxmlNode): boolean {
  if (node.kind !== 'paragraph' || node.children.length !== 1) return false;
  const properties = node.children[0]!;
  if (!is(properties, 'pPr') || properties.attributes.length || properties.children.length !== 2)
    return false;
  const style = properties.children.find((child) => is(child, 'pStyle'));
  const borders = properties.children.find((child) => is(child, 'pBdr'));
  return (
    !!style &&
    emptyProperty(style, 'pStyle', { val: SCIOBOT_HORIZONTAL_RULE_STYLE_ID }) &&
    !!borders &&
    is(borders, 'pBdr') &&
    !borders.attributes.length &&
    borders.children.length === 2 &&
    ['bottom', 'between'].every((side) =>
      borders.children.some((child) =>
        emptyProperty(child, side, { val: 'single', sz: '6', space: '1', color: 'auto' })
      )
    )
  );
}

/** Validate every reserved marker before enabling the style for a matching immutable layout. */
export function horizontalRuleProjectionStyles(pkg: OoxmlPackage): readonly string[] {
  if (horizontalRuleStyleState(pkg) !== 'present') return [];
  let valid = true;
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') {
      const properties = node.children.find((child) => is(child, 'pPr'));
      if (
        properties &&
        properties.children.some(
          (child) => is(child, 'pStyle') && attr(child, 'val') === SCIOBOT_HORIZONTAL_RULE_STYLE_ID
        ) &&
        !isHorizontalRuleParagraph(node)
      )
        valid = false;
    }
    node.children.forEach(visit);
  };
  for (const part of pkg.parts.values()) visit(part.root);
  return valid ? [SCIOBOT_HORIZONTAL_RULE_STYLE_ID] : [];
}

function fresh(node: OoxmlNode, next: () => string): OoxmlNode {
  if (node.kind === 'textValue') return { ...node, id: next() };
  return {
    ...node,
    id: next(),
    children: node.children.map((child) => fresh(child, next)),
  } as OoxmlElement;
}

/** Prepare before room bootstrap. This immutable helper creates no live edit or history entry. */
export function withHorizontalRuleStyle(pkg: OoxmlPackage): OoxmlPackage | null {
  const state = horizontalRuleStyleState(pkg);
  if (state === 'present') return pkg;
  if (state === 'conflict') return null;
  const existing = stylesPartOf(pkg);
  const main = pkg.mainDocumentPart;
  const directory = main.slice(0, main.lastIndexOf('/') + 1);
  const occupied = new Set([...pkg.parts.keys(), ...pkg.partBytes.keys()].map(partNameKey));
  let filename = 'styles.xml';
  for (let suffix = 1; occupied.has(partNameKey(directory + filename)); suffix++)
    filename = `sciobot-rule-styles-${suffix}.xml`;
  const name = existing?.name ?? directory + filename;
  const parsed = readOoxmlPart(`<w:styles xmlns:w="${WML_NAMESPACE_URI}">${STYLE}</w:styles>`, {
    name,
    contentType: CONTENT_TYPE,
  });
  if (!parsed.ok) return null;
  if (existing) {
    const inserted = insertChildren(existing, existing.root.id, existing.root.children.length, [
      fresh(parsed.part.root.children[0]!, createNodeIdAllocator(existing)),
    ]);
    return inserted.ok ? withPart(pkg, inserted.part) : null;
  }
  const created = withNewPart(pkg, name, parsed.part.root as OoxmlElement, CONTENT_TYPE);
  if (!created.parts.has(name)) return null;
  const related = withRelationship(created, main, RELATIONSHIP, filename);
  return related.ok ? related.pkg : null;
}
