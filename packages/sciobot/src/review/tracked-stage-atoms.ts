import type { InlineImageRuntime } from '../native/inline-image-runtime';
import {
  atomSourceLink,
  stagedAtomKey,
  stagedAtomWidth,
  stageAtomLink,
  atomLinkSkeletons,
  stageAtomLinkSkeletons,
} from './tracked-atom-links';
import {
  AUTHORABLE_RUN_PROPERTIES,
  collectRevisionSites,
  findNode,
  paragraphOffsetIndex,
  authoredProperties,
  contentControlPropertiesOf,
  contentControlsIn,
  parentNodeOf,
  prepareTrackedRunOrigins,
  propertyContainer,
  runsCovering,
  withPart,
} from '@docx-editor.dev/core/store';
import { blockTag, readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import { buildBlockIndex } from '../blocks/block-index';
import type { CanonicalEdit } from '../diff/plan';
import { refuseWrite } from '../diff/fragment-seeds';
import {
  inlineCodeCharacterStyles,
  inlineCodeRunChanges,
  hyperlinkCharacterStyle,
  withInlineCodeStyle,
} from '../native/inline-code';
import {
  stagedParagraphCoordinates,
  type StagedTextCoordinates,
} from './tracked-stage-coordinates';
import type { TrackedStageWriter, TrackedStageAttribution } from './tracked-stage-writer';

type Atoms = Extract<CanonicalEdit, { kind: 'inline-atoms' }>;

/** Actual native operations only: detached admission and observed publication use the same writer. */
export function stageTrackedAtoms(
  writer: TrackedStageWriter,
  edit: Atoms,
  revision: TrackedStageAttribution,
  initialSites: ReadonlySet<string>,
  independentSites: Set<string>,
  mapping: StagedTextCoordinates,
  rememberScope: (scope: string) => void,
  images?: InlineImageRuntime
): StagedTextCoordinates {
  const pkg = writer.currentPackage(),
    index = buildBlockIndex(pkg),
    payloads = readPayloadStore(pkg);
  if (payloads.status === 'invalid')
    refuseWrite('The inserted atom payload store is invalid.', 'invalid-document');
  if (mapping.text().slice(edit.offset, edit.offset + edit.before.length) !== edit.before)
    refuseWrite('The retained atom insertion text no longer matches.', 'stale-revision');
  const controls = edit.atoms.map((atom) => {
    if (atom.kind === 'drawing') {
      if (!images)
        refuseWrite('Tracked inline images need resource ownership.', 'unsupported-operation');
      return atom;
    }
    if (
      index.byId.has(atom.payload.id) ||
      (payloads.status === 'ready' && payloads.entries.has(atom.payload.id))
    )
      refuseWrite('Inserted application block IDs must be new.', 'stale-revision');
    return atom;
  });
  // Offsets include earlier fallback widths. Prepare the unmodified skeleton once.
  let width = 0;
  const ranges = controls.map((atom) => {
    const at = mapping.point(atom.offset - width, 'right');
    width += stagedAtomWidth(atom);
    return { start: at, end: at };
  });
  for (const { start, end } of atomLinkSkeletons(edit)) ranges.push(...mapping.slices(start, end));
  const part = writer.currentPackage().parts.get(edit.partName)!;
  const paragraph = findNode(part, edit.paragraphId);
  if (paragraph?.kind !== 'paragraph')
    refuseWrite('The atom paragraph is missing.', 'stale-revision');
  const offsets = paragraphOffsetIndex(paragraph);
  const bornRanges = collectRevisionSites(part).flatMap((site) => {
    const span = offsets.spanOf(site.node);
    return site.node.kind === 'revisionInsert' && !initialSites.has(site.node.id) && span
      ? [span]
      : [];
  });
  // New proposal text has no historical source run. Preserve provenance only for original
  // text; imported and previously pending revisions retain the native refusal rules.
  const originalRanges = ranges.filter(
    (range) => !bornRanges.some((span) => range.start >= span.start && range.end <= span.end)
  );
  writer.edit((working) => {
    const prepared = prepareTrackedRunOrigins(
      working.parts.get(edit.partName)!,
      edit.paragraphId,
      originalRanges
    );
    if (prepared.scope) rememberScope(prepared.scope);
    return withPart(working, prepared.part);
  });
  let current = stageAtomLinkSkeletons(writer, edit, revision, initialSites, mapping);
  const inserted = new Set<string>();
  for (const atom of controls) {
    const sourceLinkId = atomSourceLink(writer, edit, atom, inserted, current);
    const offset = current.point(atom.offset, 'right');
    const reservation =
      atom.kind === 'drawing'
        ? images!.reservation(atom, edit.paragraphId)
        : {
            op: 'insertInlineContentControl' as const,
            paragraphId: edit.paragraphId,
            offset,
            text: atom.fallback,
            tag: blockTag(atom.payload.id),
            alias: atom.label,
          };
    const tag = reservation.tag;
    writer.apply(edit.partName, { ...reservation, offset, revision });
    let part = writer.currentPackage().parts.get(edit.partName)!;
    const control = contentControlsIn(part.root).filter(
      ({ node }) => contentControlPropertiesOf(node).tag === tag
    );
    if (control.length !== 1)
      refuseWrite('The tracked atom needs one native control.', 'apply-failed');
    const native = control[0]!.node,
      wrapper = parentNodeOf(part, native.id);
    if (wrapper?.kind !== 'revisionInsert' || wrapper.children.length !== 1)
      refuseWrite('The tracked atom needs independent native ownership.', 'unsupported-operation');
    stageAtomLink(writer, edit, atom, native.id, sourceLinkId, offset);
    part = writer.currentPackage().parts.get(edit.partName)!;
    independentSites.add(wrapper.id);
    if (atom.kind === 'inline-control' && atom.fallback.length) {
      const run = runsCovering(part, edit.paragraphId, offset, offset + atom.fallback.length)[0];
      const inherited = authoredProperties(
        propertyContainer(run, 'runProperties', 'rPr'),
        AUTHORABLE_RUN_PROPERTIES
      ).filter((property) => !['b', 'i', 'strike'].includes(property.localName));
      writer.apply(edit.partName, {
        op: 'setRunProperties',
        paragraphId: edit.paragraphId,
        start: offset,
        end: offset + atom.fallback.length,
        properties: [
          ...inherited,
          ...atom.marks
            .filter((mark) => mark !== 'code')
            .map((mark) => ({
              localName: mark === 'bold' ? 'b' : mark === 'italic' ? 'i' : 'strike',
              attributes: {},
            })),
        ],
      });
      part = writer.currentPackage().parts.get(edit.partName)!;
      const code = inlineCodeRunChanges(
        part,
        {
          paragraphId: edit.paragraphId,
          start: offset,
          end: offset + atom.fallback.length,
          code: atom.marks.includes('code'),
        },
        {
          codeStyles: inlineCodeCharacterStyles(writer.currentPackage()),
          hyperlinkStyle: hyperlinkCharacterStyle(writer.currentPackage()),
        }
      );
      if (!code.success) refuseWrite(code.error, code.code);
      for (const change of code.changes) {
        if (change.after && atom.marks.includes('code'))
          writer.edit((working) => {
            const styled = withInlineCodeStyle(working, [change.after!]);
            if (!styled)
              refuseWrite('The inserted atom code style is unavailable.', 'unsupported-operation');
            return styled;
          });
        writer.apply(edit.partName, {
          op: 'setRunProperties',
          paragraphId: edit.paragraphId,
          start: change.start,
          end: change.end,
          properties: change.properties,
          characterStyle: change.after,
          ...(change.targetRunIds ? { targetRunIds: change.targetRunIds } : {}),
        });
      }
    }
    if (atom.kind === 'inline-control')
      writer.edit((working) => {
        const stored = withPayloadRecord(working, atom.payload, atom.label);
        if (!stored.ok)
          refuseWrite('The tracked atom payload could not be stored.', 'apply-failed');
        return stored.pkg;
      });
    if (atom.kind === 'inline-control')
      writer.apply(edit.partName, {
        op: 'setContentControlProperties',
        controlId: native.id,
        lock: 'sdtContentLocked',
      });
    inserted.add(stagedAtomKey(atom));
    current = stagedParagraphCoordinates(writer, edit.partName, edit.paragraphId, initialSites);
  }
  return current;
}
