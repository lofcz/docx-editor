import {
  applyTreeOp,
  ensureHyperlinkRelationship,
  findNode,
  paragraphOffsetIndex,
  replaceNode,
  withPart,
  type OoxmlElement,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { compileInlineTarget } from '../diff/inline-target';
import type { CanonicalEdit } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
type AtomEdit = Extract<CanonicalEdit, { kind: 'inline-atoms' }>;
interface LinkRange {
  readonly start: number;
  readonly end: number;
  readonly target:
    | { readonly kind: 'bookmark'; readonly anchor: string }
    | { readonly kind: 'external'; readonly url: string };
  readonly title: string | null;
  readonly original: OoxmlElement | null;
}
/** Link ranges own their complete native controls, with the original AST boundary as authority. */
export function retainedAtomLinks(pkg: OoxmlPackage, part: OoxmlPart, edit: AtomEdit) {
  const compiled = compileInlineTarget(edit.ast);
  if (!compiled.success) refuseWrite(compiled.error, compiled.code);
  const links = compiled.target.hyperlinks.filter((link) =>
    edit.atoms.some((atom) => atom.hyperlink?.path.join('/') === link.path.join('/'))
  );
  const paragraph = findNode(part, edit.paragraphId);
  if (paragraph?.kind !== 'paragraph')
    refuseWrite('The linked atom paragraph no longer exists.', 'stale-revision');
  const index = paragraphOffsetIndex(paragraph);
  const ranges: LinkRange[] = [];
  const unlink: TreeDocOp[] = [];
  for (const link of links) {
    let before = edit.offset,
      after = edit.offset;
    let start = -1,
      end = -1,
      skeletonStart = -1,
      skeletonEnd = -1;
    for (const token of compiled.target.tokens) {
      const atom =
        token.kind === 'atom'
          ? edit.atoms.find((atom) => atom.path.join('/') === token.path.join('/'))
          : undefined;
      const width =
        token.kind === 'text'
          ? token.value.length
          : atom?.kind === 'inline-control'
            ? atom.fallback.length
            : atom?.modelLength;
      if (width === undefined) refuseWrite('A linked atom has no native width.', 'invalid-input');
      const skeletonWidth = token.kind === 'text' ? token.value.length : 0;
      if (token.hyperlink?.join('/') === link.path.join('/')) {
        if (start < 0) {
          start = after;
          skeletonStart = before;
        }
        end = after + width;
        skeletonEnd = before + skeletonWidth;
      }
      before += skeletonWidth;
      after += width;
    }
    if (start < 0 || end <= start)
      refuseWrite('A linked atom needs nonempty native content.', 'unsupported-operation');
    let original: OoxmlElement | null = null;
    for (const child of paragraph.children) {
      if (child.kind !== 'hyperlink') continue;
      const span = index.spanOf(child);
      if (!span || span.end <= skeletonStart || span.start >= skeletonEnd) continue;
      const attr = (name: string) => child.attributes.find((a) => a.localName === name)?.value;
      const target = link.target;
      const matches =
        target.kind === 'bookmark'
          ? attr('anchor') === target.anchor
          : pkg.relationships
              .get(part.name)
              ?.some(
                (relationship) =>
                  relationship.id === attr('id') && relationship.rawTarget === target.url
              ) ||
            (target.kind === 'external' &&
              ensureHyperlinkRelationship(pkg, target.url, part.name)?.relationshipId ===
                attr('id'));
      if (
        original ||
        span.start !== skeletonStart ||
        span.end !== skeletonEnd ||
        !matches ||
        (attr('tooltip') ?? null) !== link.title
      )
        refuseWrite(
          'The retained hyperlink no longer matches its owned atom range.',
          'stale-revision'
        );
      original = child;
      unlink.push({ op: 'removeHyperlink', linkId: child.id });
    }
    ranges.push({ start, end, target: link.target, title: link.title, original });
  }
  return {
    unlink,
    complete(pkg: OoxmlPackage): OoxmlPackage {
      let working = pkg;
      for (const range of ranges) {
        let target: { relationshipId: string } | { anchor: string };
        if (range.target.kind === 'bookmark') target = { anchor: range.target.anchor };
        else {
          const made = ensureHyperlinkRelationship(working, range.target.url, part.name);
          if (!made)
            refuseWrite(
              'The native engine refused the linked atom target.',
              'unsupported-operation'
            );
          working = made.pkg;
          target = { relationshipId: made.relationshipId };
        }
        const before = working.parts.get(part.name)!;
        const wrapped = applyTreeOp(before, {
          op: 'insertHyperlink',
          paragraphId: edit.paragraphId,
          start: range.start,
          end: range.end,
          ...target,
          ...(range.title !== null ? { tooltip: range.title } : {}),
        });
        if (!wrapped.ok)
          refuseWrite(`Linked atom wrapping refused: ${wrapped.reason}.`, 'apply-failed');
        let next = wrapped.part;
        if (range.original) {
          const paragraph = findNode(next, edit.paragraphId);
          if (paragraph?.kind !== 'paragraph')
            refuseWrite('The linked paragraph disappeared.', 'apply-failed');
          const offsets = paragraphOffsetIndex(paragraph);
          const link = paragraph.children.find(
            (node) =>
              node.kind === 'hyperlink' &&
              offsets.spanOf(node)?.start === range.start &&
              offsets.spanOf(node)?.end === range.end
          );
          if (!link || link.kind !== 'hyperlink')
            refuseWrite('The linked atom wrapper was not allocated.', 'apply-failed');
          // Preserve the observed native link and all its imported metadata. Only its owned children change.
          const replaced = replaceNode(next, link.id, {
            ...range.original,
            attributes: range.original.attributes.map((attribute) =>
              attribute.localName === 'id'
                ? {
                    ...attribute,
                    value: 'relationshipId' in target ? target.relationshipId : attribute.value,
                  }
                : attribute
            ),
            children: link.children,
          } as OoxmlElement);
          if (!replaced.ok)
            refuseWrite('The retained native link identity could not be restored.', 'apply-failed');
          next = replaced.part;
        }
        working = withPart(working, next);
      }
      return working;
    },
  };
}
