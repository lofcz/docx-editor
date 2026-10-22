import { withInsertedParagraphPropertyPresence } from './inserted-paragraph-property-presence.ts';
// Insert complete NEW paragraphs without splitting or rewriting an existing paragraph.
import {
  createMarkContainerIdAllocator,
  createNodeIdAllocator,
  findNode,
  insertChildren,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlParagraphNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { actorScopedSeed } from '../package/actor-scoped-ids.ts';
import { usedParaIds, w14PrefixInScopeAt, withFreshParaIds } from '../package/para-id.ts';
import { buildingBlockBodyRefusal } from './building-block-safety.ts';
import { MAX_FRAGMENT_INSERT_BLOCKS, withRequiredNamespaceBindings } from './tree-op-fragment.ts';
import {
  cloneWithNewIds,
  contentControlAncestorsOf,
  fromEdit,
  isTemporaryControl,
  TEXT_DEPS,
} from './tree-op-nodes.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { build, revisionAttributes } from './tree-op-tracked-builders.ts';
import {
  invalidRevisionAttribution,
  type RevisionAttributionInput,
  type TreeOpRejection,
  type TreeOpResult,
} from './tree-op-types.ts';

/** Insert self-contained new paragraphs at a native sibling boundary. @public */
export interface InsertParagraphsOp {
  readonly op: 'insertParagraphs';
  readonly parentId: string;
  /** Existing following child; null appends before a terminal body section property. */
  readonly beforeId: string | null;
  /** New seeds only. Resource references, existing revisions and section marks refuse. */
  readonly paragraphs: readonly OoxmlParagraphNode[];
  /** Track both the inserted content and its paragraph marks with one native address. */
  readonly revision?: RevisionAttributionInput;
}

const CONTAINERS = new Set([
  'body',
  'tc',
  'sdtContent',
  'hdr',
  'ftr',
  'footnote',
  'endnote',
  'txbxContent',
]);
const wml = (node: OoxmlNode, name: string) =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;

export function validateInsertParagraphs(
  part: OoxmlPart,
  op: InsertParagraphsOp
): TreeOpRejection | null {
  const parent = typeof op.parentId === 'string' ? findNode(part, op.parentId) : null;
  if (!parent || parent.kind === 'textValue') return 'unknown-block';
  if (contentControlAncestorsOf(part, parent.id).some(isTemporaryControl)) return 'unsupported';
  if (parent.namespaceUri !== WML_NAMESPACE_URI || !CONTAINERS.has(parent.localName))
    return 'unsupported';
  if (op.beforeId !== null && !parent.children.some((node) => node.id === op.beforeId))
    return 'unknown-block';
  if (!Array.isArray(op.paragraphs) || op.paragraphs.length === 0) return 'fragment-invalid-block';
  if (op.paragraphs.length > MAX_FRAGMENT_INSERT_BLOCKS) return 'fragment-resource-budget';
  if (op.revision !== undefined && invalidRevisionAttribution(op.revision)) return 'invalidArgs';
  const refusal = buildingBlockBodyRefusal(op.paragraphs);
  if (refusal) return refusal;
  const sectionMark = (node: OoxmlNode): boolean =>
    wml(node, 'sectPr') || (node.kind !== 'textValue' && node.children.some(sectionMark));
  for (const paragraph of op.paragraphs) {
    if (paragraph.kind !== 'paragraph' || !wml(paragraph, 'p')) return 'fragment-invalid-block';
    if (sectionMark(paragraph)) return 'unsupported';
  }
  return null;
}

/** Stamp detached clones, so every node reaches collaboration as one fresh insertion. */
function trackedParagraph(
  paragraph: OoxmlParagraphNode,
  address: string,
  revision: RevisionAttributionInput,
  mint: () => string,
  markMint: () => string
): OoxmlParagraphNode {
  const properties = paragraph.children.find((node) => node.kind === 'paragraphProperties');
  const oldRPr = properties?.children.find((node) => wml(node, 'rPr'));
  const mark = withInsertedParagraphPropertyPresence(
    build(mint(), 'generic', 'ins', revisionAttributes(address, revision), []),
    properties,
    oldRPr
  );
  const rPr = {
    ...(oldRPr ?? build(markMint(), 'runProperties', 'rPr', [], [])),
    children: [mark, ...(oldRPr?.children ?? [])],
  } as OoxmlElement;
  const pPr = {
    ...(properties ?? build(markMint(), 'paragraphProperties', 'pPr', [], [])),
    children: [...(properties?.children.filter((node) => !wml(node, 'rPr')) ?? []), rPr],
  } as OoxmlElement;
  const content = paragraph.children.filter((node) => node.kind !== 'paragraphProperties');
  return {
    ...paragraph,
    children: [
      pPr,
      ...(content.length
        ? [build(mint(), 'revisionInsert', 'ins', revisionAttributes(address, revision), content)]
        : []),
    ],
  } as OoxmlParagraphNode;
}

export function applyInsertParagraphs(
  part: OoxmlPart,
  op: InsertParagraphsOp,
  options?: EditOptions
): TreeOpResult {
  const parent = findNode(part, op.parentId) as OoxmlElement;
  const mint = createNodeIdAllocator(part, 'paste');
  const markMint = createMarkContainerIdAllocator(part);
  const paraIds = new Set(usedParaIds(part.root));
  const counter = { value: 0 };
  const prefix = w14PrefixInScopeAt(part, parent);
  const address = op.revision
    ? (options?.trackedRevisionIds?.mint() ?? nextRevisionId(part)())
    : null;
  const paragraphs = op.paragraphs.map((paragraph, index) => {
    const clone = withFreshParaIds(
      cloneWithNewIds(paragraph, mint),
      paraIds,
      actorScopedSeed(`${op.parentId}:${op.beforeId}:${index}`),
      counter,
      prefix
    ) as OoxmlParagraphNode;
    return op.revision && address !== null
      ? trackedParagraph(clone, address, op.revision, mint, markMint)
      : clone;
  });
  let at =
    op.beforeId === null
      ? parent.children.findIndex((node) => wml(node, 'sectPr'))
      : parent.children.findIndex((node) => node.id === op.beforeId);
  if (at < 0) at = parent.children.length;
  return fromEdit(
    insertChildren(
      withRequiredNamespaceBindings(part, paragraphs),
      parent.id,
      at,
      paragraphs,
      options
    ),
    {
      dirty: [parent.id],
      created: paragraphs.map((paragraph) => paragraph.id),
      deleted: [],
      dependencyKeys: TEXT_DEPS,
      impact: 'flow-structural',
    }
  );
}
