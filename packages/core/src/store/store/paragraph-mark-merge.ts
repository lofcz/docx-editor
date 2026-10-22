import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlParagraphNode,
} from '../package/ooxml-tree.ts';
import { tableChildren } from './revision-table-children.ts';
import { isInertMarker } from './revision-marker-content.ts';

const isContentControl = (node: OoxmlNode): boolean =>
  node.kind === 'contentControl' ||
  (node.kind === 'generic' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'sdt');

const properties = (n: OoxmlNode) =>
  n.kind !== 'textValue' && n.namespaceUri === WML_NAMESPACE_URI && n.localName === 'pPr';

/** Word joins a paragraph before a table into its first cell's first paragraph. */
export function tableParagraphMergeTarget(
  table: OoxmlNode,
  removed?: ReadonlySet<string>
): OoxmlElement | undefined {
  if (table.kind !== 'table') return undefined;
  const row = tableChildren(table, 'tableRow').find((n) => !removed?.has(n.id));
  const cell = row && tableChildren(row, 'tableCell').find((n) => !removed?.has(n.id));
  const first = cell?.children.find(
    (n) =>
      n.kind === 'paragraph' ||
      n.kind === 'table' ||
      (n.kind !== 'textValue' && isContentControl(n))
  );
  return first?.kind === 'paragraph' ? first : undefined;
}

/** Paragraph marks whose removal can merge into the next sibling block. */
export function paragraphMergeSources(
  children: readonly OoxmlNode[],
  removed?: ReadonlySet<string>
): ReadonlySet<string> {
  const sources = new Set<string>();
  let nextAcceptsText = false;
  for (let index = children.length - 1; index >= 0; index--) {
    const child = children[index]!;
    if (removed?.has(child.id)) continue;
    if (child.kind === 'paragraph') {
      if (nextAcceptsText) sources.add(child.id);
      nextAcceptsText = true;
    } else if (child.kind === 'table') {
      nextAcceptsText = tableParagraphMergeTarget(child, removed) !== undefined;
    } else if (child.kind !== 'textValue' && isContentControl(child)) {
      nextAcceptsText = false;
    }
  }
  return sources;
}

function prepend(paragraph: OoxmlElement, content: readonly OoxmlNode[]): OoxmlElement {
  return {
    ...paragraph,
    children: [
      ...paragraph.children.filter(properties),
      ...content,
      ...paragraph.children.filter((n) => !properties(n)),
    ],
  } as OoxmlElement;
}
function replaceIn(node: OoxmlNode, target: OoxmlElement): OoxmlNode {
  if (node.id === target.id) return target;
  if (node.kind === 'textValue') return node;
  const children = node.children.map((n) => replaceIn(n, target));
  return children.some((n, i) => n !== node.children[i])
    ? ({ ...node, children } as OoxmlElement)
    : node;
}

/** Empty-mark cleanup cannot discard a remaining section or independent property revision. */
function canRemoveEmptyParagraph(node: OoxmlNode): boolean {
  if (node.kind !== 'paragraph' || node.children.some((child) => !properties(child))) return false;
  const structure = (child: OoxmlNode): boolean =>
    child.kind !== 'textValue' &&
    ((child.namespaceUri === WML_NAMESPACE_URI &&
      (child.localName === 'sectPr' ||
        child.localName.endsWith('Change') ||
        ['ins', 'del', 'moveFrom', 'moveTo'].includes(child.localName))) ||
      child.children.some(structure));
  return !node.children.some(structure);
}

function lastBlock(children: readonly OoxmlNode[]): OoxmlNode | undefined {
  for (let index = children.length - 1; index >= 0; index--) {
    const child = children[index]!;
    if (
      child.kind === 'paragraph' ||
      child.kind === 'table' ||
      (child.kind !== 'textValue' && isContentControl(child))
    )
      return child;
  }
  return undefined;
}

/** Merge after rebuilding: a removed table/row cannot receive or swallow carried text. */
export function mergeMarkedParagraphs(
  children: OoxmlNode[],
  marks: ReadonlySet<string>,
  mint?: () => string,
  retain?: (paragraph: OoxmlParagraphNode, forward: boolean) => OoxmlParagraphNode
): OoxmlNode[] {
  if (!children.some((n) => marks.has(n.id))) return children;
  const sources = paragraphMergeSources(children);
  const hasSurvivingParagraph = children.some(
    (node) => node.kind === 'paragraph' && !marks.has(node.id)
  );
  const out: OoxmlNode[] = [];
  let carried: OoxmlNode[] = [];
  for (const node of children) {
    if (node.kind === 'paragraph') {
      if (marks.has(node.id) && sources.has(node.id)) {
        if (retain) {
          out.push(retain(node, true));
          continue;
        }
        // A dropped paragraph only adds to the carried content; the one that keeps its mark
        // is built once, so a long run of joins stays linear.
        for (const child of node.children) if (!properties(child)) carried.push(child);
        continue;
      }
      // A control boundary or the container edge can prevent a forward join. The resolved
      // empty paragraph can still disappear when another paragraph survives. Keep the sole
      // required paragraph and a paragraph after a table, plus peer runs and opaque markers.
      if (
        marks.has(node.id) &&
        !sources.has(node.id) &&
        carried.length === 0 &&
        (lastBlock(out)?.kind === 'paragraph' ||
          (hasSurvivingParagraph && lastBlock(out)?.kind !== 'table')) &&
        canRemoveEmptyParagraph(node)
      ) {
        if (retain) out.push(retain(node, false));
        continue;
      }
      out.push(carried.length ? prepend(node, carried) : node);
      carried = [];
    } else if (node.kind === 'table' && carried.length) {
      const target = tableParagraphMergeTarget(node)!;
      out.push(replaceIn(node, prepend(target, carried)));
      carried = [];
    } else if (carried.length && isInertMarker(node)) {
      // A bookmark or comment boundary between the joined paragraphs marks the join point,
      // so it goes after the carried text, not before the whole joined paragraph. Native
      // decisions mint a new identity for that moved boundary. Pure Word snapshots can reuse
      // its identity because the old parent no longer contains it in that snapshot.
      carried.push(mint ? ({ ...node, id: mint() } as OoxmlNode) : node);
    } else out.push(node);
  }
  return out;
}
