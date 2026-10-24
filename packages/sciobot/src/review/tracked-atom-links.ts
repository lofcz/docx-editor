import {
  ensureHyperlinkRelationship,
  findNode,
  hyperlinkTargetOf,
  paragraphOffsetIndex,
  parentNodeOf,
  relationshipTargetIn,
  replaceChildren,
  withPart,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { CanonicalEdit } from '../diff/plan';
import type { InsertedInlineAtom } from '../diff/inserted-atoms';
import { compileInlineTarget } from '../diff/inline-target';
import { refuseWrite } from '../diff/fragment-seeds';
import type { StagedTextCoordinates } from './tracked-stage-coordinates';
import { stageTrackedLink } from './tracked-stage-links';
import { stagedParagraphCoordinates } from './tracked-stage-coordinates';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

type Control = Extract<InsertedInlineAtom, { kind: 'inline-control' }>;
type Atoms = Extract<CanonicalEdit, { kind: 'inline-atoms' }>;
function matches(pkg: OoxmlPackage, part: OoxmlPart, node: OoxmlNode, atom: Control): boolean {
  if (!atom.hyperlink || node.kind !== 'hyperlink') return false;
  const actual = hyperlinkTargetOf(node, (id) => relationshipTargetIn(pkg, part.name, id));
  const target = atom.hyperlink.target;
  return (
    (target.kind === 'external'
      ? actual.kind === 'external' && actual.authored === target.url
      : actual.kind === 'internal' && actual.anchor === target.anchor) &&
    (actual.tooltip ?? null) === atom.hyperlink.title
  );
}

/** Bind the requested AST link's existing skeleton before this atom changes native offsets. */
export function atomSourceLink(
  writer: TrackedStageWriter,
  edit: Atoms,
  atom: Control,
  inserted: ReadonlySet<string>,
  coordinates: StagedTextCoordinates
): string | null {
  if (!atom.hyperlink) return null;
  const compiled = compileInlineTarget(edit.ast);
  if (!compiled.success) refuseWrite(compiled.error, compiled.code);
  let offset = edit.offset,
    start = -1,
    end = -1;
  for (const token of compiled.target.tokens) {
    const bound =
      token.kind === 'atom'
        ? edit.atoms.find((candidate) => candidate.path.join('/') === token.path.join('/'))
        : undefined;
    const width =
      token.kind === 'text'
        ? token.value.length
        : bound?.kind === 'inline-control' && inserted.has(bound.payload.id)
          ? bound.fallback.length
          : 0;
    if (token.hyperlink?.join('/') === atom.hyperlink.path.join('/')) {
      if (start < 0) start = offset;
      end = offset + width;
    }
    offset += width;
  }
  if (start < 0 || end <= start) return null;
  const slices = coordinates.slices(start, end);
  if (!slices.length) return null;
  const pkg = writer.currentPackage(),
    part = pkg.parts.get(edit.partName)!;
  const paragraph = findNode(part, edit.paragraphId);
  if (paragraph?.kind !== 'paragraph')
    refuseWrite('The linked atom paragraph disappeared.', 'apply-failed');
  const index = paragraphOffsetIndex(paragraph);
  const links = paragraph.children.filter((node) => {
    const span = index.spanOf(node);
    return (
      span?.start === slices[0]!.start &&
      span.end === slices.at(-1)!.end &&
      matches(pkg, part, node, atom)
    );
  });
  if (links.length > 1)
    refuseWrite('The atom link has ambiguous native ownership.', 'ambiguous-target');
  return links[0]?.id ?? null;
}

/** Only the newly born insertion moves; source text and source link metadata retain their identities. */
export function stageAtomLink(
  writer: TrackedStageWriter,
  edit: Atoms,
  atom: Control,
  controlId: string,
  sourceLinkId: string | null,
  offset: number
): void {
  const pkg = writer.currentPackage(),
    part = pkg.parts.get(edit.partName)!;
  const insertion = parentNodeOf(part, controlId);
  if (insertion?.kind !== 'revisionInsert' || insertion.children.length !== 1)
    refuseWrite('The atom link needs one independent insertion.', 'unsupported-operation');
  const owner = parentNodeOf(part, insertion.id);
  if (owner?.kind === 'hyperlink') {
    if (!matches(pkg, part, owner, atom))
      refuseWrite('The atom inherited a different hyperlink.', 'unsupported-operation');
    return;
  }
  if (!atom.hyperlink) return;
  if (owner?.id !== edit.paragraphId || owner.kind !== 'paragraph')
    refuseWrite('A born atom link needs a native paragraph owner.', 'unsupported-operation');
  if (sourceLinkId) {
    const link = findNode(part, sourceLinkId);
    const index = paragraphOffsetIndex(owner),
      span = link && index.spanOf(link);
    if (!link || link.kind !== 'hyperlink' || !span || !matches(pkg, part, link, atom))
      refuseWrite('The original atom hyperlink changed.', 'stale-revision');
    const leading = offset + atom.fallback.length === span.start;
    const trailing = offset === span.end;
    if (!leading && !trailing)
      refuseWrite('The atom is not beside its bound hyperlink.', 'unsupported-operation');
    writer.edit((working) => {
      const current = working.parts.get(edit.partName)!;
      const moved = replaceChildren(
        current,
        owner.id,
        owner.children.flatMap((node) =>
          node.id === insertion.id
            ? []
            : node.id === link.id
              ? [
                  {
                    ...link,
                    children: leading
                      ? [insertion, ...link.children]
                      : [...link.children, insertion],
                  },
                ]
              : [node]
        )
      );
      if (!moved.ok)
        refuseWrite('The atom link could not retain its native owners.', 'apply-failed');
      return withPart(working, moved.part);
    });
    return;
  }
  let relationshipId: string | undefined;
  const target = atom.hyperlink.target;
  if (target.kind === 'external')
    writer.edit((working) => {
      const made = ensureHyperlinkRelationship(working, target.url, edit.partName);
      if (!made) refuseWrite('The native engine refuses the atom URL.', 'unsupported-operation');
      relationshipId = made.relationshipId;
      return made.pkg;
    });
  writer.apply(edit.partName, {
    op: 'insertHyperlink',
    paragraphId: edit.paragraphId,
    start: offset,
    end: offset + atom.fallback.length,
    ...(target.kind === 'external' ? { relationshipId } : { anchor: target.anchor }),
    ...(atom.hyperlink.title !== null ? { tooltip: atom.hyperlink.title } : {}),
  });
  const completed = writer.currentPackage(),
    current = completed.parts.get(edit.partName)!;
  const parent = parentNodeOf(current, insertion.id);
  if (!parent || !matches(completed, current, parent, atom))
    refuseWrite('The born atom hyperlink did not land exactly.', 'apply-failed');
}

/** Link ordinary source text with native revisions before inserting independently owned controls. */
export function atomLinkSkeletons(edit: Atoms) {
  const compiled = compileInlineTarget(edit.ast);
  if (!compiled.success) refuseWrite(compiled.error, compiled.code);
  return compiled.target.hyperlinks.flatMap((link) => {
    if (!edit.atoms.some((atom) => atom.hyperlink?.path.join('/') === link.path.join('/')))
      return [];
    let offset = edit.offset,
      start = -1,
      end = -1;
    for (const token of compiled.target.tokens) {
      const width = token.kind === 'text' ? token.value.length : 0;
      if (token.hyperlink?.join('/') === link.path.join('/')) {
        if (start < 0) start = offset;
        end = offset + width;
      }
      offset += width;
    }
    return start >= 0 && end > start ? [{ link, start, end }] : [];
  });
}
export function stageAtomLinkSkeletons(
  writer: TrackedStageWriter,
  edit: Atoms,
  revision: TrackedStageAttribution,
  initialSites: ReadonlySet<string>,
  mapping: StagedTextCoordinates
): StagedTextCoordinates {
  let coordinates = mapping;
  for (const { link, start, end } of atomLinkSkeletons(edit)) {
    const atom = edit.atoms.find(
      (atom): atom is Control =>
        atom.kind === 'inline-control' && atom.hyperlink?.path.join('/') === link.path.join('/')
    )!;
    if (atomSourceLink(writer, edit, atom, new Set(), coordinates)) continue;
    stageTrackedLink(
      writer,
      {
        kind: 'link',
        partName: edit.partName,
        paragraphId: edit.paragraphId,
        start,
        end,
        target: link.target,
        title: link.title,
      },
      revision,
      coordinates
    );
    coordinates = stagedParagraphCoordinates(
      writer,
      edit.partName,
      edit.paragraphId,
      initialSites,
      mapping.text()
    );
  }
  return coordinates;
}
