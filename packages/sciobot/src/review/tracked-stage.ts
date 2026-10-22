import { collectRevisionSites, sealTrackedRunOrigins, withPart } from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import type { CanonicalEdit, CanonicalEditPlan } from '../diff/plan';
import type { ReviewGroupReceipt } from './groups';
import { stageTrackedFormat } from './tracked-stage-format';
import { stageTrackedSiblings } from './tracked-stage-siblings';
import { stagedParagraphCoordinates } from './tracked-stage-coordinates';
import { currentStagePart } from './tracked-stage-text';
import {
  stageTrackedLink,
  linkStageParagraph,
  existingLinkStageParagraph,
  coupleLinkStageReceipts,
  stageParagraphKey as key,
} from './tracked-stage-links';
import {
  stageTrackedTable,
  trackedTableId,
  coupleTableStageReceipts,
  tableStageHasStructure,
} from './tracked-stage-table';
import { stageTrackedText, type StagedTextCoordinates } from './tracked-stage-text';
import type { TrackedStageAttribution, TrackedStageWriter } from './tracked-stage-writer';

/** Ordered lowering shared by detached admission and the actual native transaction. */
export function stageTrackedPlan(
  writer: TrackedStageWriter,
  plan: CanonicalEditPlan,
  revision: TrackedStageAttribution,
  trackFormatting: boolean
): readonly ReviewGroupReceipt[] {
  const receipts: ReviewGroupReceipt[] = [];
  const formatReceipts = new Map<string, ReviewGroupReceipt[]>();
  const tableReceipts = new Map<string, ReviewGroupReceipt[]>();
  const structuralTables = new Set<string>();
  const initial = writer.currentPackage();
  const unselectedSites = new Set(
    collectRevisionSites(initial.parts.get(initial.mainDocumentPart)!).map((site) => site.node.id)
  );
  const coordinates = new Map<string, StagedTextCoordinates>();
  const linkParagraphs = new Set<string>();
  const deferredFormats: Extract<CanonicalEdit, { kind: 'format' }>[] = [];
  for (const edit of plan.operations) {
    if (
      edit.kind !== 'link' &&
      edit.kind !== 'unlink' &&
      !(edit.kind === 'format' && edit.target.kind === 'hyperlink')
    )
      continue;
    if (edit.partName !== initial.mainDocumentPart)
      refuseWrite(
        'Replacement plans must belong to the native body story.',
        'unsupported-operation'
      );
    let paragraphId: string;
    if (edit.kind === 'link' || edit.kind === 'unlink')
      paragraphId = linkStageParagraph(initial.parts.get(initial.mainDocumentPart)!, edit);
    else if (edit.kind === 'format' && edit.target.kind === 'hyperlink')
      paragraphId = existingLinkStageParagraph(
        initial.parts.get(initial.mainDocumentPart)!,
        edit.target.hyperlinkId
      );
    else continue;
    linkParagraphs.add(key(edit.partName, paragraphId));
  }
  const format = (edit: Extract<CanonicalEdit, { kind: 'format' }>) => {
    if (edit.target.kind === 'hyperlink') {
      const paragraphId = existingLinkStageParagraph(
        currentStagePart(writer, edit.partName),
        edit.target.hyperlinkId
      );
      const paragraph = key(edit.partName, paragraphId);
      if (!coordinates.has(paragraph))
        coordinates.set(
          paragraph,
          stagedParagraphCoordinates(writer, edit.partName, paragraphId, unselectedSites)
        );
    }

    const identity = key(
      edit.partName,
      edit.target.kind === 'hyperlink' ? edit.target.hyperlinkId : edit.target.paragraphId
    );
    const entries = formatReceipts.get(identity) ?? [];
    entries.push(
      ...stageTrackedFormat(
        writer,
        edit,
        revision,
        edit.target.kind === 'run'
          ? coordinates.get(key(edit.partName, edit.target.paragraphId))
          : undefined,
        unselectedSites
      )
    );
    formatReceipts.set(identity, entries);
  };
  for (let at = 0; at < plan.operations.length; ) {
    const edit = plan.operations[at]!;
    if (edit.partName !== writer.currentPackage().mainDocumentPart)
      refuseWrite(
        'Replacement plans must belong to the native body story.',
        'unsupported-operation'
      );
    if (edit.kind === 'text') {
      const entries: Extract<CanonicalEdit, { kind: 'text' }>[] = [edit];
      at++;
      while (at < plan.operations.length) {
        const next = plan.operations[at]!;
        if (
          next.kind !== 'text' ||
          next.partName !== edit.partName ||
          next.paragraphId !== edit.paragraphId
        )
          break;
        entries.push(next);
        at++;
      }
      const paragraph = key(edit.partName, edit.paragraphId);
      if (coordinates.has(paragraph))
        refuseWrite('A paragraph needs one ordered text stage.', 'ambiguous-target');
      coordinates.set(
        paragraph,
        stageTrackedText(writer, plan, entries, revision, linkParagraphs.has(paragraph))
      );
      continue;
    }
    if (edit.kind === 'format') {
      if (!trackFormatting)
        refuseWrite(
          'Formatting tracking is disabled for this document or editor.',
          'unsupported-operation'
        );
      if (
        edit.target.kind === 'run' &&
        linkParagraphs.has(key(edit.partName, edit.target.paragraphId))
      )
        deferredFormats.push(edit);
      else format(edit);
    } else if (edit.kind === 'link' || edit.kind === 'unlink') {
      const paragraphId = linkStageParagraph(currentStagePart(writer, edit.partName), edit);
      const identity = key(edit.partName, paragraphId);
      let mapping = coordinates.get(identity);
      if (!mapping) {
        mapping = stagedParagraphCoordinates(writer, edit.partName, paragraphId, unselectedSites);
        coordinates.set(identity, mapping);
      }
      stageTrackedLink(writer, edit, revision, mapping);
    } else if (edit.kind === 'siblings')
      receipts.push(...stageTrackedSiblings(writer, edit, revision));
    else if (
      ['table', 'insert-row', 'insert-column', 'split-cell', 'table-geometry'].includes(edit.kind)
    ) {
      const tableEdit = edit as Extract<
        CanonicalEdit,
        { kind: 'table' | 'insert-row' | 'insert-column' | 'split-cell' | 'table-geometry' }
      >;
      const identity = key(edit.partName, trackedTableId(tableEdit));
      if (tableStageHasStructure(tableEdit)) structuralTables.add(identity);
      const entries = tableReceipts.get(identity) ?? [];
      entries.push(...stageTrackedTable(writer, tableEdit, revision, trackFormatting));
      tableReceipts.set(identity, entries);
    } else
      refuseWrite(
        `Replacement ${edit.kind} needs its tracked staged executor.`,
        'unsupported-operation'
      );
    at++;
  }
  for (const edit of deferredFormats) format(edit);
  for (const entries of [...formatReceipts.values(), ...tableReceipts.values()]) {
    if (!entries.length) continue;
    receipts.push({
      partName: entries[0]!.partName,
      addresses: entries.flatMap((entry) => entry.addresses),
    });
  }
  for (const [identity, mapping] of coordinates) {
    if (!mapping.runOriginScope) continue;
    const [partName, paragraphId] = JSON.parse(identity) as [string, string];
    const scope = mapping.runOriginScope;
    writer.edit((pkg) =>
      withPart(
        pkg,
        sealTrackedRunOrigins(
          pkg.parts.get(partName)!,
          paragraphId,
          scope,
          initial.parts.get(partName)!
        )
      )
    );
  }
  const linkedReceipts = coupleLinkStageReceipts(writer, receipts, linkParagraphs, unselectedSites);
  return coupleTableStageReceipts(writer, linkedReceipts, structuralTables, unselectedSites);
}
