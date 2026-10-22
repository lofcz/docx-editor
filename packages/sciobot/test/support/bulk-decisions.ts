import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  collectRevisionSites,
  paragraphTextOf,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import { readPropertyReviewStore } from '../../src/review/property-record-store';
import { executeAtomic, proposedListening, proposedVisual } from './atomic-execution';
import { executeProperties } from './property-execution';
import { proposeGroupText } from './review-groups';

export const bulkSource = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
export function proposeMixedBulk(editor: DocxEditorInstance, betweenWaves?: () => void) {
  proposeGroupText(editor, 'The cat sat.', 'The car sat.');
  betweenWaves?.();
  proposeGroupText(editor, 'A dog ran.', 'A dog runs.');
  const properties = executeProperties(editor);
  if (!properties.success) throw new Error(properties.error);
  const atomic = executeAtomic(editor, [proposedVisual, proposedListening]);
  if (!atomic.success) throw new Error(atomic.error);
}
export function payloadSnapshot(pkg: OoxmlPackage) {
  const read = readPayloadStore(pkg);
  if (read.status !== 'ready') throw new Error('Missing payloads');
  return [...read.entries].map(([id, entry]) => [id, entry.raw, entry.label]);
}
export function proposalSnapshot(pkg: OoxmlPackage) {
  const read = readPropertyReviewStore(pkg);
  if (read.status !== 'ready') throw new Error('Missing proposals');
  return [...read.records].map(([id, entry]) => [id, entry.raw]);
}
/** Independent oracle: native all-revision acceptance and the explicit proposed payload values. */
export function acceptedMixedPackage(
  pkg: OoxmlPackage,
  selectedSites?: ReadonlyMap<string, readonly string[]>
) {
  let next = pkg;
  for (const part of pkg.parts.values()) {
    if (!collectRevisionSites(part).length) continue;
    const accepted = applyTreeOp(part, {
      op: 'acceptAllRevisions',
      ...(selectedSites ? { siteNodeIds: [...(selectedSites.get(part.name) ?? [])] } : {}),
    });
    if (!accepted.ok) throw new Error(accepted.reason);
    next = withPart(next, accepted.part);
  }
  const owner = buildBlockIndex(pkg).byId.get('activity')!.payload;
  for (const [payload, label] of [
    [
      { ...owner, attributes: { ...owner.attributes, name: 'New activity', duration: 15 } },
      'New activity — 15 minutes',
    ],
    [proposedVisual, 'New visual fallback'],
    [proposedListening, 'New audio fallback'],
  ] as const) {
    const written = withPayloadRecord(next, payload, label);
    if (!written.ok) throw new Error(written.reason);
    next = written.pkg;
  }
  return next;
}

/** Actual manual input between proposal waves; it must survive every bulk decision. */
export function typeBulkSentence(editor: DocxEditorInstance) {
  const surface = editor.surface!,
    id = surface.session
      .paragraphIds()
      .find((id) => paragraphTextOf(surface.session.part(), id) === 'Keep tail');
  if (!id) throw new Error('Missing manual sentence paragraph');
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  surface.insertPlainText('Manual sentence. ');
  surface.flushPendingInput();
  return id;
}

/** Independent oracle: native rejection retains live committed payload records and peer content. */
export function rejectedMixedPackage(
  pkg: OoxmlPackage,
  selectedSites?: ReadonlyMap<string, readonly string[]>
) {
  let next = pkg;
  for (const part of pkg.parts.values()) {
    if (!collectRevisionSites(part).length) continue;
    const rejected = applyTreeOp(part, {
      op: 'rejectAllRevisions',
      ...(selectedSites ? { siteNodeIds: [...(selectedSites.get(part.name) ?? [])] } : {}),
    });
    if (!rejected.ok) throw new Error(rejected.reason);
    next = withPart(next, rejected.part);
  }
  return next;
}
