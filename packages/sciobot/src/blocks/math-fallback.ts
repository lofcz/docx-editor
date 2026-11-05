import type { OoxmlNode, OoxmlParagraphNode } from '@docx-editor.dev/core/store';

/** Keep the existing text identity when a fallback line changes. */
function lineWithIdentity(template: OoxmlNode, previous: OoxmlNode | undefined): OoxmlNode {
  if (!previous || template.kind !== previous.kind) return template;
  if (template.kind === 'textValue') return { ...template, id: previous.id };
  if (previous.kind === 'textValue') return template;
  return {
    ...template,
    id: previous.id,
    children: template.children.map((child, index) =>
      lineWithIdentity(child, previous.children[index])
    ),
  } as OoxmlNode;
}

/** Display mode moves the same fallback runs between a control and its paragraphs. */
export function stableMathFallback(
  content: OoxmlNode,
  paragraphs: readonly OoxmlNode[],
  inline: boolean
): readonly OoxmlNode[] {
  if (content.kind === 'textValue') throw Error('The equation content is invalid.');
  const oldParagraphs: OoxmlParagraphNode[] = [];
  for (const child of content.children) if (child.kind === 'paragraph') oldParagraphs.push(child);
  const oldRuns = oldParagraphs.length
    ? oldParagraphs.flatMap((paragraph) =>
        paragraph.children.filter((child) => child.kind === 'run')
      )
    : content.children.filter((child) => child.kind === 'run');
  let run = 0;
  const converted = paragraphs.map((paragraph, index) => {
    if (paragraph.kind !== 'paragraph') throw Error('The fallback paragraph is invalid.');
    const previous = oldParagraphs[index];
    return {
      ...paragraph,
      ...(previous ? { id: previous.id, attributes: previous.attributes } : {}),
      children: paragraph.children.map((child) =>
        child.kind === 'run' ? lineWithIdentity(child, oldRuns[run++]) : child
      ),
    } as OoxmlParagraphNode;
  });
  return inline ? converted[0]!.children : converted;
}
