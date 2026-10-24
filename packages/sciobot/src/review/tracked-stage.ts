import { stageTrackedAtoms } from './tracked-stage-atoms';
import { stageTrackedInlineRemoval } from './tracked-stage-inline-removal';
import { insertedAtomsOfEdit } from '../diff/inserted-cell-atoms';
import { createDirectListContext } from '../ai/direct-lists';
import { stageTrackedList } from './tracked-stage-lists';
import {
  collectRevisionSites,
  prepareTrackedRunOrigins,
  sealTrackedRunOrigins,
  withPart,
} from '@docx-editor.dev/core/store';
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
  if (
    plan.operations.some((edit) => {
      const atoms = insertedAtomsOfEdit(edit);
      if (!atoms.length || edit.kind === 'inline-atoms') return false;
      if (edit.kind !== 'siblings') return true;
      return atoms.some(
        (atom) =>
          atom.kind !== 'inline-control' ||
          atom.location.kind !== 'new-paragraph' ||
          !['paragraph', 'heading', 'blockquote'].includes(
            edit.inserted[atom.location.path[0]!]?.type ?? ''
          )
      );
    })
  )
    refuseWrite(
      'Inserted inline atoms need native asset and payload reconciliation.',
      'unsupported-operation'
    );
  const listContext = createDirectListContext();
  const receipts: ReviewGroupReceipt[] = [];
  const formatReceipts = new Map<string, ReviewGroupReceipt[]>();
  const tableReceipts = new Map<string, ReviewGroupReceipt[]>();
  const structuralTables = new Set<string>();
  const initial = writer.currentPackage();
  const unselectedSites = new Set(
    collectRevisionSites(initial.parts.get(initial.mainDocumentPart)!).map((site) => site.node.id)
  );
  const coordinates = new Map<string, StagedTextCoordinates>();
  const scopes = new Map<string, Set<string>>();
  const atomSites = new Set<string>();
  const rememberScope = (identity: string, mapping: StagedTextCoordinates) => {
    if (!mapping.runOriginScope) return;
    const values = scopes.get(identity) ?? new Set<string>();
    values.add(mapping.runOriginScope);
    scopes.set(identity, values);
  };
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
  // Pure link plans can split a source run even when no text stage runs.
  // Prepare every link range together before any replacement or formatting changes it.
  const textParagraphs = new Set(
    plan.operations.flatMap((edit) =>
      edit.kind === 'text' ? [key(edit.partName, edit.paragraphId)] : []
    )
  );
  const pureLinks = new Map<string, Extract<CanonicalEdit, { kind: 'link' }>[]>();
  for (const edit of plan.operations) {
    if (edit.kind !== 'link') continue;
    const identity = key(edit.partName, edit.paragraphId);
    if (textParagraphs.has(identity)) continue;
    const ranges = pureLinks.get(identity) ?? [];
    ranges.push(edit);
    pureLinks.set(identity, ranges);
  }
  const preparePureLinks = (identity: string) => {
    const edits = pureLinks.get(identity);
    if (!edits) return;
    const first = edits[0]!;
    const mapping =
      coordinates.get(identity) ??
      stagedParagraphCoordinates(writer, first.partName, first.paragraphId, unselectedSites);
    let scope: string | null = null;
    writer.edit((pkg) => {
      const prepared = prepareTrackedRunOrigins(
        pkg.parts.get(first.partName)!,
        first.paragraphId,
        edits.flatMap((edit) => mapping.slices(edit.start, edit.end))
      );
      scope = prepared.scope;
      return withPart(pkg, prepared.part);
    });
    coordinates.set(identity, {
      ...stagedParagraphCoordinates(writer, first.partName, first.paragraphId, unselectedSites),
      ...(scope ? { runOriginScope: scope } : {}),
    });
    rememberScope(identity, coordinates.get(identity)!);
    pureLinks.delete(identity);
  };
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
        unselectedSites,
        (scope) => {
          if (edit.target.kind !== 'run') return;
          const identity = key(edit.partName, edit.target.paragraphId);
          const values = scopes.get(identity) ?? new Set<string>();
          values.add(scope);
          scopes.set(identity, values);
        }
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
      const mapping = coordinates.get(paragraph);
      coordinates.set(
        paragraph,
        stageTrackedText(
          writer,
          plan,
          entries,
          revision,
          mapping ? { coordinates: mapping, initialSites: unselectedSites } : undefined
        )
      );
      rememberScope(paragraph, coordinates.get(paragraph)!);
      continue;
    }
    if (edit.kind === 'inline-remove') {
      const beforeSites = new Set(
        collectRevisionSites(currentStagePart(writer, edit.partName)).map((site) => site.node.id)
      );
      stageTrackedInlineRemoval(writer, edit, revision);
      for (const site of collectRevisionSites(currentStagePart(writer, edit.partName)))
        if (!beforeSites.has(site.node.id)) atomSites.add(site.node.id);
      // Only earlier native deletions are absent from this operation's proposed coordinates.
      coordinates.set(
        key(edit.partName, edit.paragraphId),
        stagedParagraphCoordinates(writer, edit.partName, edit.paragraphId, unselectedSites)
      );
      at++;
      continue;
    } else if (edit.kind === 'inline-atoms') {
      const identity = key(edit.partName, edit.paragraphId);
      const mapping =
        coordinates.get(identity) ??
        stagedParagraphCoordinates(writer, edit.partName, edit.paragraphId, unselectedSites);
      coordinates.set(
        identity,
        stageTrackedAtoms(writer, edit, revision, unselectedSites, atomSites, mapping, (scope) => {
          const values = scopes.get(identity) ?? new Set<string>();
          values.add(scope);
          scopes.set(identity, values);
        })
      );
    } else if (edit.kind === 'format') {
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
      preparePureLinks(identity);
      let mapping = coordinates.get(identity);
      if (!mapping) {
        mapping = stagedParagraphCoordinates(writer, edit.partName, paragraphId, unselectedSites);
        coordinates.set(identity, mapping);
      }
      stageTrackedLink(writer, edit, revision, mapping);
    } else if (edit.kind === 'list') {
      if (!trackFormatting)
        refuseWrite(
          'Formatting tracking is disabled for this document or editor.',
          'unsupported-operation'
        );
      receipts.push(...stageTrackedList(writer, plan, edit, revision, listContext));
    } else if (edit.kind === 'siblings')
      receipts.push(...stageTrackedSiblings(writer, edit, revision, listContext));
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
      entries.push(...stageTrackedTable(writer, tableEdit, revision, trackFormatting, listContext));
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
  for (const [identity, values] of scopes) {
    const [partName, paragraphId] = JSON.parse(identity) as [string, string];
    for (const scope of values)
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
  const linkedReceipts = coupleLinkStageReceipts(
    writer,
    receipts,
    linkParagraphs,
    unselectedSites,
    atomSites
  );
  return coupleTableStageReceipts(writer, linkedReceipts, structuralTables, unselectedSites);
}
