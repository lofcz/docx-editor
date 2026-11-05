import {
  bodyStoryRoot,
  contentControlPropertiesOf,
  contentControlsIn,
  createNodeIdAllocator,
  escapeXml,
  parentNodeOf,
  readOoxmlPart,
  replaceChildren,
  usedParaIds,
  w14RootPrefix,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { withFreshIds } from './fresh-ids';
import { stableMathFallback } from './math-fallback';
import { BLOCK_MATH, INLINE_MATH, mathWordFallback, type MathPayload } from './math-codec';
import { blockTag } from './payload-opc-codec';

/** Reparent a canonical equation control. Adjacent runs remain the original nodes. */
export function convertMathControl(
  pkg: OoxmlPackage,
  partName: string,
  payload: MathPayload
): OoxmlPackage {
  const part = pkg.parts.get(partName);
  if (!part) throw Error('The equation story is missing.');
  const control = contentControlsIn(part.root).find(
    ({ node }) => contentControlPropertiesOf(node).tag === blockTag(payload.id)
  )?.node;
  const content = control?.children.find((child) => child.kind === 'contentControlContent');
  const parent = control && parentNodeOf(part, control.id);
  if (!control || !content || !parent) throw Error('The equation is missing.');
  const inline = parent.kind === 'paragraph';
  if (!inline && !['body', 'tableCell', 'contentControlContent'].includes(parent.kind))
    throw Error('This equation cannot change display in its container.');
  const fallback = mathWordFallback(payload);
  const lines = payload.type === INLINE_MATH ? [fallback.inlineText] : fallback.paragraphs;
  const xml =
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>` +
    lines
      .map((line) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`)
      .join('') +
    '</w:body></w:document>';
  const parsed = readOoxmlPart(xml, { name: '/math-fallback.xml', contentType: part.contentType });
  const body = parsed.ok && bodyStoryRoot(parsed.part);
  if (!body || body.kind === 'textValue') throw Error('The equation fallback is invalid.');
  const allocate = createNodeIdAllocator(part, 'new');
  const identity = {
    prefix: w14RootPrefix(part.root),
    used: new Set(usedParaIds(part.root)),
    seed: `math:${payload.id}`,
  };
  const fresh = (node: OoxmlNode) => withFreshIds(node, allocate, identity);
  const paragraphs = body.children.map(fresh);
  const first = paragraphs[0]!;
  if (first.kind !== 'paragraph') throw Error('The equation fallback has no paragraph.');
  const converted: OoxmlNode = {
    ...control,
    children: control.children.map((child) =>
      child.id === content.id
        ? {
            ...content,
            children: stableMathFallback(content, paragraphs, payload.type === INLINE_MATH),
          }
        : child
    ),
  };
  let target = parent;
  let replacements: OoxmlNode[];
  let removedId = control.id;
  if (inline && payload.type === BLOCK_MATH) {
    const container = parentNodeOf(part, parent.id);
    if (!container || !['body', 'tableCell', 'contentControlContent'].includes(container.kind))
      throw Error('This paragraph cannot contain a display equation.');
    const index = parent.children.findIndex((child) => child.id === control.id);
    const properties = parent.children.filter((child) => child.kind === 'paragraphProperties');
    const before = parent.children
      .slice(0, index)
      .filter((child) => child.kind !== 'paragraphProperties');
    const after = parent.children
      .slice(index + 1)
      .filter((child) => child.kind !== 'paragraphProperties');
    replacements = [];
    if (before.length) replacements.push({ ...parent, children: [...properties, ...before] });
    replacements.push(converted);
    if (after.length) {
      // Keep the original paragraph identity on one side, never both.
      const tail = before.length
        ? fresh({ ...body.children[0]!, children: [] } as OoxmlNode)
        : parent;
      if (tail.kind !== 'paragraph') throw Error('The equation tail is invalid.');
      replacements.push({
        ...tail,
        children: [...(before.length ? properties.map(fresh) : properties), ...after],
      } as OoxmlNode);
    }
    target = container;
    removedId = parent.id;
  } else if (!inline && payload.type === INLINE_MATH) {
    replacements = [{ ...first, children: [converted] }];
  } else {
    replacements = [converted];
  }
  const children = target.children.flatMap((child) =>
    child.id === removedId ? replacements : [child]
  );
  const replaced = replaceChildren(part, target.id, children);
  if (!replaced.ok) throw Error('The equation conversion violates document structure.');
  return withPart(pkg, replaced.part);
}
