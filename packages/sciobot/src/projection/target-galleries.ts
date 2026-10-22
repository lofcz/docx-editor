// Application-owned gallery transform; no editor runtime imports.
import type { RootContent } from 'mdast';
import type { Plugin } from 'unified';

interface GalleryImageData {
  url: string;
  caption?: string;
}

/**
 * remark plugin: group consecutive image-only paragraphs into a single
 * `<image_gallery />` MDX block (an mdxJsxFlowElement node).
 *
 * This runs at markdown deserialization time inside `MarkdownPlugin` — so
 * every `![alt](url)` the AI emits (or any markdown image authored by hand)
 * is rendered as a compact, aspect-preserving thumbnail with a PhotoSwipe
 * lightbox, instead of a full-width `<img>` blow-up.
 *
 * What it transforms
 *
 *   ![One](u1)              →   <image_gallery layout="row-fit"
 *                                  images={[{ "url": "u1" }]} />
 *
 *   ![One](u1)              →   <image_gallery layout="row-fit"
 *   ![Two](u2)                      images={[{ "url": "u1" },
 *                                            { "url": "u2" }]} />
 *
 *   Text before.            →   Text before.
 *   ![One](u1)                  <image_gallery .../>
 *   Text after.                 Text after.
 *
 * What it leaves alone
 *
 *   • Inline images inside prose (e.g. `Here is ![cat](u) the cat.`).
 *   • Existing `<image_gallery />` blocks (they are mdxJsxFlowElement
 *     nodes, not `image`/`paragraph` nodes, so they pass through).
 *
 * Where it descends
 *
 *   • The root document and every node with a `children` array — most
 *     importantly `mdxJsxFlowElement` containers like `<activity>` and
 *     `<callout>`, so an image rendered between list items inside an
 *     activity is grouped into a gallery just the same as one at the top
 *     level.
 *
 * Why a remark plugin (and not regex on the markdown string)
 *
 *   • Robust against arbitrary MDX nesting — we never have to reinvent a
 *     parser to know whether an `![](…)` lives inside a code fence, a
 *     quote, an activity block, etc.
 *   • Idempotent — re-running on already-transformed markdown is a no-op
 *     because galleries are not `image` or `paragraph` nodes.
 *   • Authoring-friendly — the persisted markdown stays as `![]()` until
 *     the editor saves, so the AI prompts and human edits both work with
 *     the natural markdown form.
 */
export const remarkImageGalleries: Plugin = () => (tree) => {
  transformContainer(tree as unknown as NodeWithChildren);
};

interface NodeWithChildren {
  children: unknown[];
}

interface MdastNode {
  type: string;
  children?: unknown[];
  [key: string]: unknown;
}

/**
 * mdast node types whose `children` carry BLOCK content (BlockContent /
 * FlowContent in the mdast spec). We recurse only into these — never into
 * phrasing parents like `paragraph`, `heading`, `tableCell`, `link`, etc.
 *
 * Why this matters: each `![](url)` in source is parsed as
 *
 *   paragraph { children: [image] }
 *
 * If we descended into the paragraph and rewrote the `image` child to an
 * `<image_gallery />` mdxJsxFlowElement, the gallery would end up *inside*
 * a paragraph (paragraph > image_gallery), which is structurally invalid —
 * paragraphs can only host phrasing content. Plate's mdast→slate converter
 * then collapses the paragraph and silently drops all but the last block
 * child, so every gallery except the last one vanishes from the document.
 *
 * Instead, we treat paragraphs as terminal nodes: their contents are
 * inspected by `extractImagesIfImageOnly` to decide whether the *whole*
 * paragraph should be replaced (at the parent block level) with a single
 * gallery carrying every image the paragraph held.
 */
const BLOCK_CONTAINER_TYPES = new Set<string>([
  'root',
  'blockquote',
  'list',
  'listItem',
  'mdxJsxFlowElement',
  'footnoteDefinition',
  'containerDirective',
]);

function transformContainer(node: NodeWithChildren): void {
  if (!Array.isArray(node.children)) return;

  // Depth-first: recurse into block-level containers so nested transforms
  // (e.g. inside `<activity>` or a `listItem` continuation paragraph) are
  // done before we regroup this level.
  for (const child of node.children) {
    if (
      child &&
      typeof child === 'object' &&
      BLOCK_CONTAINER_TYPES.has((child as MdastNode).type)
    ) {
      transformContainer(child as unknown as NodeWithChildren);
    }
  }

  // Sweep the immediate children, collapsing runs of image-only nodes
  // (bare `image` siblings and/or whole image-only paragraphs) into one
  // gallery at this level.
  const next: unknown[] = [];
  let i = 0;
  while (i < node.children.length) {
    const images = extractImagesIfImageOnly(node.children[i]);
    if (images.length === 0) {
      next.push(node.children[i]);
      i += 1;
      continue;
    }
    let j = i + 1;
    while (j < node.children.length) {
      const more = extractImagesIfImageOnly(node.children[j]);
      if (more.length === 0) break;
      images.push(...more);
      j += 1;
    }
    next.push(createGalleryNode(images));
    i = j;
  }
  node.children = next;
}

/**
 * Returns the images carried by `node` IFF the node is an image-only block
 * (a bare `image` node, or a paragraph whose only meaningful children are
 * images / image-wrapped links). Otherwise returns an empty array — which
 * the caller treats as "leave this node alone".
 */
function extractImagesIfImageOnly(node: unknown): GalleryImageData[] {
  if (!node || typeof node !== 'object') return [];
  const n = node as MdastNode;

  if (n.type === 'image') {
    const img = imageFromMdast(n);
    return img ? [img] : [];
  }

  if (n.type !== 'paragraph' || !Array.isArray(n.children)) return [];

  const out: GalleryImageData[] = [];
  for (const raw of n.children) {
    if (!raw || typeof raw !== 'object') return [];
    const child = raw as MdastNode;

    if (child.type === 'image') {
      const img = imageFromMdast(child);
      if (!img) return [];
      out.push(img);
      continue;
    }

    // `[![alt](src)](href)` — markdown for an image wrapped in a link.
    // Strip the link, keep the inner image.
    if (child.type === 'link' && Array.isArray(child.children)) {
      const inner = extractSingleImageFromLink(child.children);
      if (inner) {
        out.push(inner);
        continue;
      }
      return [];
    }

    // Whitespace-only text nodes between images are fine — they just
    // separate consecutive `![](…)` on the same paragraph line.
    if (child.type === 'text' && typeof child.value === 'string' && child.value.trim() === '') {
      continue;
    }

    return [];
  }

  return out;
}

function extractSingleImageFromLink(linkChildren: unknown[]): GalleryImageData | null {
  let found: GalleryImageData | null = null;
  for (const raw of linkChildren) {
    if (!raw || typeof raw !== 'object') return null;
    const c = raw as MdastNode;
    if (c.type === 'image') {
      if (found) return null; // multiple images inside the link — bail
      found = imageFromMdast(c);
      if (!found) return null;
      continue;
    }
    if (c.type === 'text' && typeof c.value === 'string' && c.value.trim() === '') {
      continue;
    }
    return null;
  }
  return found;
}

function imageFromMdast(node: MdastNode): GalleryImageData | null {
  const url = typeof node.url === 'string' ? node.url.trim() : '';
  if (!url) return null;
  // The Markdown alt text (`![caption](url)`) becomes the visible caption —
  // this is the channel the drafting LLM uses to label images.
  const caption = typeof node.alt === 'string' ? node.alt.trim() : '';
  return caption ? { url, caption } : { url };
}

function createGalleryNode(images: GalleryImageData[]): unknown {
  return {
    type: 'mdxJsxFlowElement',
    name: 'image_gallery',
    attributes: [
      { type: 'mdxJsxAttribute', name: 'layout', value: 'row-fit' },
      {
        type: 'mdxJsxAttribute',
        name: 'images',
        value: {
          type: 'mdxJsxAttributeValueExpression',
          value: JSON.stringify(images),
        },
      },
    ],
    children: [],
  } satisfies MdastNode;
}

// Silences "RootContent imported but unused" by tying it to the public type.
// remark plugins receive `Root` which contains `RootContent[]`; the rest of
// this module uses structural typing because mdxJsxFlowElement is provided
// by `mdast-util-mdx-jsx` and we don't want a hard dep on it.
export type _RootContent = RootContent;
