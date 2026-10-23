// Known failing investigation fixture. Run explicitly with bun test on this path.
import { expect, test } from 'bun:test';
import { createTwoEditorFixture } from '../../../../../packages/sciobot/test/support/collaboration';
import { docxFixture } from '../../../../../packages/sciobot/test/support/document';
import { table } from '../../../../../packages/sciobot/test/support/target-table-fixture';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { prepareNativeDocument } from '../../../../../packages/sciobot/src/native/prepare-document';
import { readEditorRevision } from '../../../../../packages/sciobot/src/projection/revision-cache';
import { readWorkingView } from '../../../../../packages/sciobot/src/review/working-view';
import { planTargetDocument } from '../../../../../packages/sciobot/src/diff/plan-target';
import { applyDirectPlan } from '../../../../../packages/sciobot/src/ai/apply-direct';
import { buildBlockIndex } from '../../../../../packages/sciobot/src/blocks/block-index';
for (const topology of ['row', 'column'])
  test(`peer edits inside a born ${topology} survive the inserting actor undo`, async () => {
    const prepared = prepareNativeDocument(
      docxFixture(
        table([
          ['A', 'B'],
          ['Keep', 'Peer'],
        ])
      )
    );
    if (!prepared.success) throw Error(prepared.error);
    const pair = await createTwoEditorFixture({ document: prepared.document, clientIds: [41, 73] });
    try {
      const [alice, bob] = pair.editors;
      const context = { documentId: 'born-cell-probe', epoch: 'native' };
      const revision = readEditorRevision(alice!, context.documentId, context.epoch)!;
      const source = await readWorkingView(alice!.surface!.session, revision);
      if (!source.success) throw Error(source.error);
      const target =
        topology === 'row'
          ? '| A | B |\n| - | - |\n| Keep | Peer |\n| $x$ | Tail |'
          : '| A | New | B |\n| - | - | - |\n| Keep | Tail $x$ | Peer |';
      const planned = planTargetDocument(
        {
          revision,
          pkg: alice!.surface!.session.currentPackage(),
          projection: source.view.projection,
        },
        target
      );
      if (!planned.success) throw Error(planned.error);
      const applied = applyDirectPlan(alice!, planned.plan, context);
      if (!applied.success) throw Error(applied.error);
      await pair.settle();
      const paragraph = bob!
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob!.surface!.session.part(), id)?.startsWith('Tail'))!;
      expect(paragraph).toBeDefined();
      const end = paragraphTextOf(bob!.surface!.session.part(), paragraph)!.length;
      const edit = bob!.surface!.applyAutomationOps(
        () => [{ op: 'insertText', paragraphId: paragraph, offset: end, text: ' peerborn' }],
        { kind: 'body' },
        [],
        false,
        true
      );
      expect(edit.rejected).toBe(false);
      await pair.settle();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        const texts = editor
          .surface!.session.paragraphIds()
          .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
        expect(texts.join('\n')).toContain('peerborn');
        expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
      }
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  });
