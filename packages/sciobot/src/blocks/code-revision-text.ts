import {
  contentControlContentNodeOf,
  hiddenRetiredParagraphs,
  isJoinedParagraph,
  WML_NAMESPACE_URI,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { DEFAULT_PAYLOAD_LIMITS } from './payload-schema';

export type CodeSourceView = 'plain' | 'committed' | 'working';
const wml = (node: OoxmlNode, name: string) =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;

/** Read code text from native runs and paragraph marks, without resolving their revisions. */
export function codeRevisionText(control: OoxmlNode, view: CodeSourceView) {
  const content = contentControlContentNodeOf(control);
  if (!content || !content.children.length || content.children.length > 10000) return null;
  if (!content.children.every((paragraph) => paragraph.kind === 'paragraph')) return null;
  const hidden = hiddenRetiredParagraphs(content.children);
  const paragraphs = content.children.filter((paragraph) => !hidden.has(paragraph.id));
  let visited = 0,
    characters = 0;
  const runs = (nodes: readonly OoxmlNode[], depth = 0): string | null => {
    if (depth > 16) return null;
    let value = '';
    for (const node of nodes) {
      if (++visited > 100000) return null;
      if (node.kind === 'revisionInsert' || node.kind === 'revisionDelete') {
        if (view === 'plain') return null;
        const text = runs(node.children, depth + 1);
        if (text === null) return null;
        if (node.kind !== (view === 'committed' ? 'revisionInsert' : 'revisionDelete'))
          value += text;
        continue;
      }
      if (node.kind !== 'run') return null;
      for (const child of node.children) {
        if (++visited > 100000) return null;
        if (child.kind === 'runProperties') {
          if (
            child.children.some((property) =>
              ['rPrChange', 'vanish', 'webHidden', 'specVanish'].includes(property.localName)
            )
          )
            return null;
        } else if (
          (child.kind === 'text' || (view !== 'plain' && wml(child, 'delText'))) &&
          child.children.every((text) => text.kind === 'textValue')
        ) {
          for (const text of child.children) {
            if (text.kind !== 'textValue') return null;
            characters += text.value.length;
            if (characters > DEFAULT_PAYLOAD_LIMITS.maxBytes * 2) return null;
            value += text.value;
          }
        } else if (wml(child, 'tab') && !child.children.length) value += '\t';
        else return null;
      }
    }
    return value;
  };
  const lines: string[] = [];
  const breaks: boolean[] = [];
  for (const paragraph of paragraphs) {
    let join = isJoinedParagraph(paragraph);
    for (const properties of paragraph.children.filter(
      (child) => child.kind === 'paragraphProperties'
    )) {
      for (const property of properties.children) {
        if (['sectPr', 'pPrChange'].includes(property.localName)) return null;
        if (wml(property, 'rPr')) {
          // Shared history can retain the empty native mark-property owner after a decision.
          if (!property.children.length) continue;
          if (view === 'plain') return null;
          for (const mark of property.children) {
            if (
              (!wml(mark, 'ins') && !wml(mark, 'del')) ||
              mark.kind === 'textValue' ||
              mark.children.length
            )
              return null;
            if (mark.localName === (view === 'committed' ? 'ins' : 'del')) join = true;
          }
        }
      }
    }
    const text = runs(paragraph.children.filter((child) => child.kind !== 'paragraphProperties'));
    if (text === null) return null;
    lines.push(text);
    breaks.push(!join);
  }
  const value = lines
    .map((line, index) => line + (index < lines.length - 1 && breaks[index] ? '\n' : ''))
    .join('');
  if (
    value.length > DEFAULT_PAYLOAD_LIMITS.maxBytes ||
    new TextEncoder().encode(value).byteLength > DEFAULT_PAYLOAD_LIMITS.maxBytes
  )
    return null;
  return { content, paragraphs, lines, breaks, value };
}
