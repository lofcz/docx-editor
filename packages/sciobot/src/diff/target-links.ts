import {
  findNode,
  hyperlinkTargetOf,
  paragraphOffsetIndex,
  relationshipTargetIn,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import type { CompiledInlineTarget } from './inline-target';
import type { CanonicalEdit } from './plan';
import { planHyperlinkFormat } from './run-format';

export interface LinkStructureSource {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  /** The paragraph's part after this plan's text edits, as an executor will hold it. */
  readonly part: OoxmlPart;
  readonly paragraphId: string;
  /** Native offset of the first model character. */
  readonly base: number;
}

/** Target link spans in post-text native offsets, in document order. */
function targetSpans(target: CompiledInlineTarget, base: number) {
  const spans = new Map<string, { start: number; end: number }>();
  let offset = base;
  for (const token of target.tokens) {
    if (token.kind !== 'text') continue;
    const end = offset + token.value.length;
    if (token.hyperlink) {
      const key = token.hyperlink.join('/'),
        old = spans.get(key);
      spans.set(key, { start: old?.start ?? offset, end });
    }
    offset = end;
  }
  return target.hyperlinks.map((link) => ({ link, span: spans.get(link.path.join('/')) }));
}

/**
 * Plan link insertion, removal, boundary changes, and retargeting in one retained paragraph.
 * Wrappers whose exact span survives keep their identity and relationship; every other wrapper is
 * removed (its runs stay) and target links over new spans are inserted. Runs after text edits.
 */
export function planLinkStructure(
  source: LinkStructureSource,
  target: CompiledInlineTarget
): { success: true; operations: readonly CanonicalEdit[] } | NativeEditFailure {
  const paragraph = findNode(source.part, source.paragraphId);
  if (paragraph?.kind !== 'paragraph')
    return { success: false, code: 'invalid-document', error: 'The linked paragraph is missing.' };
  const wrappers: OoxmlNode[] = [],
    stack: OoxmlNode[] = [paragraph];
  while (stack.length) {
    const node = stack.pop()!;
    if (node.kind === 'hyperlink') wrappers.push(node);
    else if (node.kind !== 'textValue') stack.push(...node.children);
  }
  const offsets = paragraphOffsetIndex(paragraph);
  const current = wrappers.map((node) => ({ node, span: offsets.spanOf(node) }));
  const wanted = targetSpans(target, source.base);
  if (wanted.some((entry) => !entry.span || entry.span.start === entry.span.end))
    return {
      success: false,
      code: 'unsupported-operation',
      error: 'A target hyperlink needs nonempty text.',
    };
  const pkg = withPart(source.pkg, source.part),
    kept = new Set<OoxmlNode>(),
    unlink: CanonicalEdit[] = [],
    link: CanonicalEdit[] = [],
    retarget: CanonicalEdit[] = [];
  for (const { link: next, span } of wanted) {
    const match = current.find(
      (entry) =>
        !kept.has(entry.node) && entry.span?.start === span!.start && entry.span.end === span!.end
    );
    const tooltipRemoved =
      match &&
      next.title === null &&
      hyperlinkTargetOf(match.node as Extract<OoxmlNode, { kind: 'hyperlink' }>, (id) =>
        relationshipTargetIn(pkg, source.part.name, id)
      ).tooltip !== undefined;
    if (match && !tooltipRemoved) {
      kept.add(match.node);
      const result = planHyperlinkFormat(
        { revision: source.revision, pkg, partName: source.part.name, hyperlinkId: match.node.id },
        {
          ...(next.target.kind === 'external'
            ? { url: next.target.url }
            : { anchor: next.target.anchor }),
          ...(next.title !== null ? { tooltip: next.title } : {}),
        }
      );
      if (!result.success) return result;
      retarget.push(...result.plan.operations);
      continue;
    }
    link.push({
      kind: 'link',
      partName: source.part.name,
      paragraphId: source.paragraphId,
      start: span!.start,
      end: span!.end,
      target: next.target,
      title: next.title,
    });
  }
  for (const entry of current)
    if (!kept.has(entry.node))
      unlink.push({ kind: 'unlink', partName: source.part.name, hyperlinkId: entry.node.id });
  return { success: true, operations: [...unlink, ...link, ...retarget] };
}
