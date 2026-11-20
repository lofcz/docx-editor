import { expect, test } from 'bun:test';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { targetDocumentInput, targetParagraph } from '../../test/support/target-document';
import { applyDirectPlan, previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { buildBlockIndex } from './block-index';
import { readPresentationColumns } from './column-geometry';
import { readPayloadStore } from './payload-opc-codec';

test('deleting overview materials preserves the peer curriculum edit, history, and saved controls', async () => {
  const context = { documentId: 'overview-column-history', epoch: '1' };
  const original = await targetDocumentInput(targetParagraph('End.'));
  const target =
    '<preparation_recap>\n\n<recap_curriculum>\n\nOutcome\n\n</recap_curriculum>\n\n<recap_materials>\n\nCards\n\n</recap_materials>\n\n</preparation_recap>\n\nEnd.';
  let id = 0;
  const create = planTargetDocument(original, target, {
    allocateBlockId: () => `overview-${++id}`,
  });
  if (!create.success) throw new Error(create.error);
  const seed = previewDirectPlan(original.pkg, create.plan);
  if (!seed.success) throw new Error(seed.error);
  const pair = await createTwoEditorFixture({
    document: writeOoxmlPackage(seed.pkg),
    documentId: context.documentId,
    clientIds: [37, 53],
  });
  try {
    const [first, second] = pair.editors;
    const surface = first!.surface!;
    const columns = () =>
      pair.editors.map((editor) => {
        const pkg = editor.surface!.session.currentPackage(),
          index = buildBlockIndex(pkg);
        const owner = [...index.byId.values()].find(
          (block) => block.payload.type === 'preparation_recap'
        )!;
        return readPresentationColumns(pkg, owner, index)!.columns.map((cell) => [
          cell.cellId,
          cell.block.id,
        ]);
      });
    const originalColumns = columns();
    let revision = readEditorRevision(first!, context.documentId, context.epoch)!;
    let read = await readWorkingView(surface.session, revision);
    if (!read.success) throw new Error(read.error);
    const parts = read.view.projection.text.match(
      /<recap_(curriculum|materials)\b[\s\S]*?<\/recap_\1>/g
    )!;
    const swapped = read.view.projection.text.replace(
      parts.join('\n\n'),
      [...parts].reverse().join('\n\n')
    );
    const reorder = planTargetDocument(
      { pkg: surface.session.currentPackage(), revision, projection: read.view.projection },
      swapped
    );
    if (!reorder.success) throw new Error(reorder.error);
    const moved = applyDirectPlan(first!, reorder.plan, context);
    if (!moved.success) throw new Error(moved.error);
    await pair.settle();
    expect(columns()).toEqual(originalColumns.map((order) => [...order].reverse()));
    expect(first!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(columns()).toEqual(originalColumns);
    revision = readEditorRevision(first!, context.documentId, context.epoch)!;
    read = await readWorkingView(surface.session, revision);
    if (!read.success) throw new Error(read.error);
    const projection = read.view.projection;
    const next = projection.text.replace(/<recap_materials>[\s\S]*?<\/recap_materials>\n\n/, '');
    const planned = planTargetDocument(
      { pkg: surface.session.currentPackage(), revision, projection },
      next
    );
    if (!planned.success) throw new Error(planned.error);
    const before = buildBlockIndex(surface.session.currentPackage());
    const material = [...before.byId.values()].find(
      (block) => block.payload.type === 'recap_materials'
    )!;
    const outcome = surface.session
      .paragraphIds()
      .find((p) => paragraphTextOf(surface.session.part(), p) === 'Outcome')!;
    second!.surface!.setSelection({
      anchor: { paragraphId: outcome, offset: 7 },
      head: { paragraphId: outcome, offset: 7 },
    });
    await pair.settle();
    pair.pause();
    const applied = applyDirectPlan(first!, planned.plan, context);
    if (!applied.success) throw new Error(applied.error);
    expect(
      second!.surface!.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: outcome, offset: 7, text: ' peer' },
      ]).rejected
    ).toBe(false);
    pair.reconnect();
    await pair.settle();
    const check = async (materialPresent: boolean) => {
      for (const editor of pair.editors) {
        const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!opened.ok) throw new Error(opened.reason);
        const index = buildBlockIndex(opened.package),
          records = readPayloadStore(opened.package);
        if (records.status !== 'ready') throw new Error(records.status);
        expect(index.byId.has(material.id)).toBe(materialPresent);
        expect(records.entries.has(material.id)).toBe(materialPresent);
        expect(index.diagnostics).toEqual([]);
        expect(paragraphTextOf(editor.surface!.session.part(), outcome)).toBe('Outcome peer');
      }
    };
    await check(false);
    expect(first!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    await check(true);
    expect(first!.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    await check(false);
    const digests = await pair.savedDigests();
    expect(digests[0]).toEqual(digests[1]);
  } finally {
    pair.destroy();
  }
}, 120_000);
