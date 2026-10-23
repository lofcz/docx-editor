import type { SerializedBlockTag } from './block-tags';
import type { TargetAstNode } from './parse-target';

const escape = (value: string) =>
  value.replace(
    /[&"<>\r\n\t|]/g,
    (char) =>
      ({
        '&': '&amp;',
        '"': '&quot;',
        '<': '&lt;',
        '>': '&gt;',
        '\r': '&#13;',
        '\n': '&#10;',
        '\t': '&#9;',
        '|': '&#124;',
      })[char]!
  );
/** Literal mention syntax can sit inside an outer Markdown link without nesting Markdown links. */
export function serializeLinkedMention(key: string, value: string): SerializedBlockTag {
  let text = '<scio-mention';
  const properties: SerializedBlockTag['properties'][number][] = [];
  for (const [property, data] of [
    ['key', key],
    ['value', value],
  ] as const) {
    text += ' ';
    const start = text.length;
    text += property + '="';
    text += escape(data);
    text += '"';
    properties.push({
      name: property,
      property,
      path: [property],
      start,
      end: text.length,
      valueStart: null,
      valueEnd: null,
    });
  }
  return { tag: 'mention', text: text + ' />', properties };
}
/** Unknown attributes, expressions, duplicate fields, and children keep their opaque meaning. */
export function literalMention(node: TargetAstNode): { key: string; value: string } | null {
  if (
    !['mdxJsxTextElement', 'mdxJsxFlowElement'].includes(node.type) ||
    node.name !== 'scio-mention' ||
    !Array.isArray(node.attributes) ||
    node.attributes.length !== 2 ||
    !Array.isArray(node.children) ||
    node.children.length
  )
    return null;
  const values = new Map<string, string>();
  for (const attr of node.attributes as { type?: unknown; name?: unknown; value?: unknown }[]) {
    if (
      attr.type !== 'mdxJsxAttribute' ||
      !['key', 'value'].includes(String(attr.name)) ||
      typeof attr.value !== 'string' ||
      values.has(String(attr.name))
    )
      return null;
    values.set(String(attr.name), attr.value);
  }
  return values.has('key') && values.has('value')
    ? { key: values.get('key')!, value: values.get('value')! }
    : null;
}
