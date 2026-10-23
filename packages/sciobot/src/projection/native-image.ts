import type { Node } from 'mdast';
import type { Plugin } from 'unified';
import type { Parent } from 'unist';
import { visit } from 'unist-util-visit';
import type { TargetAstNode } from './parse-target';
/** Exact inert data syntax. Expressions and extension metadata remain opaque. */
export const remarkNativeImage: Plugin = () => (tree) => {
  visit(
    tree,
    ['mdxJsxTextElement', 'mdxJsxFlowElement'],
    (raw: Node, index: number | undefined, parent: Parent | undefined) => {
      if (!parent || index === undefined) return;
      const node = raw as unknown as TargetAstNode;
      if (
        node.name !== 'scio-inline-image' ||
        !Array.isArray(node.children) ||
        node.children.length ||
        !Array.isArray(node.attributes) ||
        node.attributes.length < 2 ||
        node.attributes.length > 3 ||
        Object.keys(node).some(
          (key) => !['type', 'name', 'children', 'attributes', 'position'].includes(key)
        )
      )
        return;
      const values = new Map<string, string>();
      for (const raw of node.attributes) {
        if (
          !raw ||
          typeof raw !== 'object' ||
          Array.isArray(raw) ||
          raw.type !== 'mdxJsxAttribute' ||
          typeof raw.name !== 'string' ||
          !['reference', 'alt', 'title'].includes(raw.name) ||
          typeof raw.value !== 'string' ||
          values.has(raw.name) ||
          Object.keys(raw).some((key) => !['type', 'name', 'value', 'position'].includes(key))
        )
          return;
        values.set(raw.name, raw.value);
      }
      if (!values.has('reference') || !values.has('alt')) return;
      const image = {
        type: 'image',
        url: values.get('reference')!,
        alt: values.get('alt')!,
        title: values.get('title') ?? null,
        ...(node.position ? { position: node.position } : {}),
      };
      parent.children[index] = (node.type === 'mdxJsxFlowElement'
        ? {
            type: 'paragraph',
            children: [image],
            ...(node.position ? { position: node.position } : {}),
          }
        : image) as unknown as Node;
    }
  );
};
