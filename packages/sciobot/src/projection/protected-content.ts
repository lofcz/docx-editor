import {
  AUTHORABLE_PARAGRAPH_PROPERTIES,
  AUTHORABLE_RUN_PROPERTIES,
  bodyStoryRoot,
  findNode,
  isFieldChrome,
  isFldSimple,
  isInstrText,
  isRangeMarkerKind,
  paragraphOffsetIndex,
  storyParagraphs,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import { sameRevision, type NativeEditFailure, type RevisionToken } from '../contracts/document';
import { createEditPlan } from '../diff/plan';
import type { MarkdownProjection } from './markdown';
import {
  outsideBodyBoundary,
  protectedRangeBoundaries,
  type ProtectedContentBoundary,
} from './protected-boundaries';
import {
  hasProtectedShell,
  nodeFingerprint,
  packageMetadataFingerprint,
  protectedPartsOf,
  type ProtectedNodeReference,
  type ProtectedParentOrder,
  type ProtectedPartReference,
  type ProtectionReason,
} from './protected-snapshots';

export type { ProtectedContentBoundary, ProtectedNativeRange } from './protected-boundaries';
export type {
  ProtectedNodeReference,
  ProtectedParentOrder,
  ProtectedPartReference,
  ProtectionReason,
} from './protected-snapshots';
export interface ProtectedContentIndex {
  readonly schemaVersion: 1;
  readonly revision: RevisionToken;
  readonly mainDocumentPart: string;
  readonly boundaries: readonly ProtectedContentBoundary[];
  readonly nodes: readonly ProtectedNodeReference[];
  readonly parts: readonly ProtectedPartReference[];
  readonly parentOrder: readonly ProtectedParentOrder[];
  readonly packageMetadata: string;
}
export type ProtectedContentResult =
  | { readonly success: true; readonly index: ProtectedContentIndex }
  | NativeEditFailure;
export type ProtectedContentCheck =
  | { readonly success: true }
  | (NativeEditFailure & {
      readonly protectedReference?: ProtectedNodeReference | ProtectedPartReference;
    });
export const PROTECTED_CONTENT_MAX_NODES = 1_000_000;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const metadataNames = new Set([
  ...AUTHORABLE_PARAGRAPH_PROPERTIES,
  ...AUTHORABLE_RUN_PROPERTIES,
  ...'numPr ilvl numId tabs tab pBdr top left bottom right between bar rPrChange pPrChange rPr pPr'.split(
    ' '
  ),
]);
const wholeProperties = new Set([
  'tableProperties',
  'tableRowProperties',
  'tableCellProperties',
  'tableGrid',
  'contentControlProperties',
  'contentControlEndProperties',
]);
interface Entry {
  node: OoxmlNode;
  parentId: string | null;
  paragraphId: string | null;
  insideBody: boolean;
  metadata: boolean;
  ordinal: number;
}
const failure = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-document'
): NativeEditFailure => ({ success: false, code, error });

/** Read immutable canonical references and Markdown boundaries for a body-text patch. */
export function collectProtectedContent(
  pkg: OoxmlPackage,
  projection: MarkdownProjection,
  current: RevisionToken
): ProtectedContentResult {
  const token = createEditPlan(current, []);
  if (!token.success) return token;
  if (!sameRevision(projection.revision, token.plan.revision))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read a current projection before collecting protected content.',
      expectedRevision: Object.freeze({ ...projection.revision }),
      actualRevision: token.plan.revision,
    };
  const part = pkg.parts.get(pkg.mainDocumentPart),
    body = part && bodyStoryRoot(part);
  if (!part || !body) return failure('The document has no canonical body.');
  const entries = new Map<string, Entry>();
  const stack = [
    {
      node: part.root as OoxmlNode,
      parentId: null as string | null,
      paragraphId: null as string | null,
      insideBody: false,
      metadata: false,
    },
  ];
  while (stack.length) {
    const next = stack.pop()!,
      node = next.node;
    if (entries.has(node.id) || entries.size >= PROTECTED_CONTENT_MAX_NODES)
      return failure(
        'The canonical body is cyclic, duplicated, or exceeds its protection scan limit.'
      );
    const entry: Entry = {
      ...next,
      ordinal: entries.size,
      insideBody: next.insideBody || node.id === body.id,
      paragraphId: node.kind === 'paragraph' ? node.id : next.paragraphId,
      metadata:
        next.metadata ||
        wholeProperties.has(node.kind) ||
        (node.kind !== 'textValue' && (node.localName === 'pPr' || node.localName === 'rPr')),
    };
    entries.set(node.id, entry);
    if (node.kind !== 'textValue')
      for (let at = node.children.length - 1; at >= 0; at--)
        stack.push({
          node: node.children[at]!,
          parentId: node.id,
          paragraphId: entry.paragraphId,
          insideBody: entry.insideBody,
          metadata: entry.metadata,
        });
  }
  const paragraphs = storyParagraphs(body).filter(
    (node): node is OoxmlParagraphNode => node.kind === 'paragraph'
  );
  const bodyIds = new Set(paragraphs.map((node) => node.id)),
    paths = new Set<string>();
  for (const paragraph of paragraphs) {
    let id: string | null = paragraph.id;
    while (id && !paths.has(id)) {
      paths.add(id);
      id = entries.get(id)?.parentId ?? null;
    }
  }
  const mapped = new Map<string, (typeof projection.sourceMap.paragraphs)[number][]>();
  for (const paragraph of projection.sourceMap.paragraphs) {
    if (
      (paragraph.markdownStart !== undefined || paragraph.markdownEnd !== undefined) &&
      (paragraph.markdownStart === undefined ||
        !Number.isSafeInteger(paragraph.markdownStart) ||
        paragraph.markdownStart < 0 ||
        paragraph.markdownEnd === undefined ||
        !Number.isSafeInteger(paragraph.markdownEnd) ||
        paragraph.markdownEnd < paragraph.markdownStart ||
        paragraph.markdownEnd > projection.text.length)
    )
      return failure('A protected projection paragraph has invalid Markdown bounds.');
    if (paragraph.sourceScope === 'body' && bodyIds.has(paragraph.paragraphId)) {
      const list = mapped.get(paragraph.paragraphId) ?? [];
      list.push(paragraph);
      mapped.set(paragraph.paragraphId, list);
    }
  }
  const sourceGroups = new Map<string, (typeof projection.sourceMap.sources)[number][]>();
  for (const source of projection.sourceMap.sources) {
    if (
      ![source.sourceStart, source.sourceEnd, source.markdownStart, source.markdownEnd].every(
        Number.isSafeInteger
      ) ||
      source.sourceStart < 0 ||
      source.sourceEnd < source.sourceStart ||
      source.markdownStart < 0 ||
      source.markdownEnd < source.markdownStart ||
      source.markdownEnd > projection.text.length
    )
      return failure('A protected projection source has invalid bounds.');
    if (source.sourceScope === 'body') {
      const list = sourceGroups.get(source.paragraphId) ?? [];
      list.push(source);
      sourceGroups.set(source.paragraphId, list);
    }
  }
  const ordered = [...mapped.entries()]
    .map(([id, ranges]) => ({
      ordinal: entries.get(id)!.ordinal,
      start: ranges.reduce((at, range) => Math.min(at, range.markdownStart ?? Infinity), Infinity),
      end: ranges.reduce((at, range) => Math.max(at, range.markdownEnd ?? -Infinity), -Infinity),
    }))
    .sort((a, b) => a.ordinal - b.ordinal);
  const blockGap = (ordinal: number): number => {
    let low = 0,
      high = ordered.length;
    while (low < high) {
      const mid = (low + high) >>> 1;
      if (ordered[mid]!.ordinal < ordinal) low = mid + 1;
      else high = mid;
    }
    const previous = ordered[low - 1],
      following = ordered[low];
    return previous && Number.isFinite(previous.end)
      ? previous.end
      : following && Number.isFinite(following.start)
        ? following.start
        : 0;
  };
  const boundaries: ProtectedContentBoundary[] = [],
    nodes: ProtectedNodeReference[] = [];
  const referenced = new Set<string>(),
    covered = new Set<string>(),
    parentIds = new Set<string>();
  const reference = (
    entry: Entry,
    kind: ProtectionReason,
    reason: string,
    mode: 'subtree' | 'shell' = 'subtree'
  ) => {
    const key = entry.node.id + '\0' + mode;
    if (referenced.has(key)) return;
    referenced.add(key);
    if (entry.parentId) parentIds.add(entry.parentId);
    nodes.push(
      Object.freeze({
        partName: part.name,
        nodeId: entry.node.id,
        parentId: entry.parentId,
        kind,
        reason,
        mode,
        fingerprint: nodeFingerprint(entry.node, mode),
      })
    );
    if (mode === 'subtree') covered.add(entry.node.id);
  };
  const boundary = (entry: Entry, start: number, end: number, reason: string) => {
    if (!entry.paragraphId) return;
    const scoped: MarkdownProjection = {
      ...projection,
      sourceMap: {
        ...projection.sourceMap,
        sources: sourceGroups.get(entry.paragraphId) ?? [],
        paragraphs: mapped.get(entry.paragraphId) ?? [],
        tables: [],
      },
    };
    boundaries.push(
      ...protectedRangeBoundaries(
        scoped,
        { paragraphId: entry.paragraphId, start, end },
        part.name,
        entry.node.id,
        reason
      )
    );
  };
  for (const paragraph of paragraphs) {
    const offsets = paragraphOffsetIndex(paragraph);
    for (const segment of offsets.segments)
      if (segment.removeNodeIds?.length) {
        const node = segment.node;
        const kind: ProtectionReason =
          isFieldChrome(node) || isFldSimple(node)
            ? 'field'
            : node.kind === 'drawing' ||
                (node.kind !== 'textValue' &&
                  ['pict', 'object', 'AlternateContent'].includes(node.localName))
              ? 'drawing'
              : 'inline-atom';
        const reason =
          kind === 'field'
            ? 'Word field'
            : kind === 'drawing'
              ? 'Word drawing'
              : 'Native inline atom';
        const entry = entries.get(node.id);
        for (const id of segment.removeNodeIds) {
          const owner = entries.get(id);
          if (owner) reference(owner, kind, reason);
        }
        if (entry) boundary(entry, segment.start, segment.end, reason);
      }
  }
  const hiddenByAncestor = new Set<string>();
  for (const entry of entries.values()) {
    const node = entry.node;
    if (entry.parentId && (covered.has(entry.parentId) || hiddenByAncestor.has(entry.parentId))) {
      hiddenByAncestor.add(node.id);
      continue;
    }
    if (covered.has(node.id)) continue;
    if (node.kind === 'textValue') continue;
    const offsets =
      entry.paragraphId && bodyIds.has(entry.paragraphId)
        ? paragraphOffsetIndex(entries.get(entry.paragraphId)!.node as OoxmlParagraphNode)
        : null;
    const span = offsets?.spanOf(node);
    let kind: ProtectionReason | null = null,
      reason = '';
    if (!entry.insideBody && !paths.has(node.id)) {
      kind = 'outside-body';
      reason = 'XML outside the writable body';
    } else if (node.kind === 'paragraph' && !mapped.has(node.id)) {
      kind = 'opaque-xml';
      reason = 'Unprojected native paragraph';
    } else if (node.namespaceUri === W && node.localName === 'sectPr') {
      kind = 'section-properties';
      reason = 'Word section properties';
    } else if (isFieldChrome(node) || isInstrText(node) || isFldSimple(node)) {
      kind = 'field';
      reason = 'Word field';
    } else if (
      isRangeMarkerKind(node.kind) ||
      node.kind === 'bookmarkStart' ||
      node.kind === 'bookmarkEnd'
    ) {
      kind = 'range-marker';
      reason = 'Native range marker';
    } else if (wholeProperties.has(node.kind)) {
      kind = 'opaque-xml';
      reason = 'Native structural properties';
    } else if (
      node.kind === 'generic' &&
      !paths.has(node.id) &&
      !(span && span.end > span.start) &&
      !(entry.metadata && node.namespaceUri === W && metadataNames.has(node.localName))
    ) {
      kind = 'opaque-xml';
      reason = 'Unprojected OOXML';
    }
    if (kind) {
      reference(entry, kind, reason);
      if (!entry.metadata && kind !== 'section-properties' && kind !== 'outside-body') {
        if (span) boundary(entry, span.start, span.end, reason);
        else if (!entry.paragraphId) {
          const at = blockGap(entry.ordinal);
          boundaries.push(
            Object.freeze({
              start: at,
              end: at,
              edges: 'interior',
              partName: part.name,
              nodeId: node.id,
              reason,
              sourceScope: 'body',
            })
          );
        }
      }
    } else if (hasProtectedShell(node))
      reference(entry, 'element-shell', 'Unprojected attributes or namespace bindings', 'shell');
  }
  for (const paragraph of projection.sourceMap.paragraphs)
    if (paragraph.sourceScope !== 'body' || !bodyIds.has(paragraph.paragraphId)) {
      const owner = [...pkg.parts].find(([, candidate]) =>
        findNode(candidate, paragraph.paragraphId)
      );
      if (!owner)
        return failure(
          `The projection names a missing native paragraph (${paragraph.paragraphId}).`
        );
      const region = outsideBodyBoundary(paragraph, owner[0]);
      if (region) boundaries.push(region);
    }
  const parentOrder: ProtectedParentOrder[] = [];
  for (const id of parentIds) {
    const parent = entries.get(id)!.node;
    if (parent.kind !== 'textValue')
      parentOrder.push(
        Object.freeze({
          partName: part.name,
          nodeId: id,
          childIds: Object.freeze(parent.children.map((node) => node.id)),
        })
      );
  }
  return {
    success: true,
    index: Object.freeze({
      schemaVersion: 1,
      revision: token.plan.revision,
      mainDocumentPart: part.name,
      boundaries: Object.freeze(
        boundaries
          .map((boundary) => Object.freeze({ ...boundary, scope: 'body-text' as const }))
          .sort((a, b) => a.start - b.start || a.end - b.end)
      ),
      nodes: Object.freeze(nodes),
      parts: protectedPartsOf(pkg),
      parentOrder: Object.freeze(parentOrder),
      packageMetadata: packageMetadataFingerprint(pkg),
    }),
  };
}

/** Validate an immutable native preview before committing body-text operations. */
export function checkProtectedContent(
  index: ProtectedContentIndex,
  candidate: OoxmlPackage,
  current: RevisionToken
): ProtectedContentCheck {
  const token = createEditPlan(current, []);
  if (!token.success) return token;
  if (index.schemaVersion !== 1)
    return failure('Invalid protected-content index.', 'invalid-input');
  if (!sameRevision(index.revision, token.plan.revision))
    return {
      success: false,
      code: 'stale-revision',
      error: 'Read protected content again before planning this body patch.',
      expectedRevision: index.revision,
      actualRevision: token.plan.revision,
    };
  const refused = (
    reference: ProtectedNodeReference | ProtectedPartReference,
    message: string
  ): ProtectedContentCheck => ({
    success: false,
    code: 'unsupported-operation',
    error: `The body patch changes protected content (${reference.partName}${'nodeId' in reference ? '#' + reference.nodeId : ''}): ${message}`,
    protectedReference: reference,
  });
  if (
    candidate.mainDocumentPart !== index.mainDocumentPart ||
    packageMetadataFingerprint(candidate) !== index.packageMetadata
  )
    return failure(
      'The body patch changes protected package relationships or content types.',
      'unsupported-operation'
    );
  const parts = protectedPartsOf(candidate);
  const partsByName = new Map(parts.map((part) => [part.partName, part]));
  for (const before of index.parts) {
    const after = partsByName.get(before.partName);
    if (!after || after.mode !== before.mode || after.fingerprint !== before.fingerprint)
      return refused(before, 'Untouched package part');
  }
  if (parts.length !== index.parts.length)
    return failure('The body patch changes protected package parts.', 'unsupported-operation');
  for (const reference of index.nodes) {
    const part = candidate.parts.get(reference.partName),
      node = part && findNode(part, reference.nodeId);
    if (!node || nodeFingerprint(node, reference.mode) !== reference.fingerprint)
      return refused(reference, reference.reason);
    const parent = reference.parentId && part && findNode(part, reference.parentId);
    if (
      reference.parentId &&
      (!parent ||
        parent.kind === 'textValue' ||
        !parent.children.some((child) => child.id === reference.nodeId))
    )
      return refused(reference, 'Protected node parent');
  }
  for (const order of index.parentOrder) {
    const part = candidate.parts.get(order.partName),
      parent = part && findNode(part, order.nodeId);
    if (!parent || parent.kind === 'textValue')
      return failure(
        `The body patch removes a protected parent (${order.nodeId}).`,
        'unsupported-operation'
      );
    const rank = new Map(order.childIds.map((id, at) => [id, at]));
    let last = -1;
    for (const node of parent.children) {
      const at = rank.get(node.id);
      if (at === undefined) continue;
      if (at < last)
        return failure(
          `The body patch reorders protected parent content (${order.nodeId}).`,
          'unsupported-operation'
        );
      last = at;
    }
  }
  return { success: true };
}
