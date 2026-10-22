import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const leadingDocument = () =>
  docxFixture(
    '<w:p><w:ins w:id="2" w:author="Reviewer" w:date="2026-10-19T09:00:00Z"><w:r><w:t>Newer</w:t></w:r></w:ins><w:del w:id="1" w:author="Reviewer" w:date="2026-10-19T09:00:00Z"><w:r><w:delText>Old</w:delText></w:r></w:del></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
  );

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const action of ['acceptReviewItem', 'rejectReviewItem'] as const)
      test(`leading native replacement ${action} preserves peer text, placement, and history; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
        const bytes = leadingDocument();
        const pair = await createTwoEditorFixture({
          document: bytes,
          clientIds,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors,
            id = alice!.surface!.session.paragraphIds()[0]!,
            tail = alice!.surface!.session.paragraphIds()[1]!;
          bob!.surface!.setSelection({
            anchor: { paragraphId: id, offset: 8 },
            head: { paragraphId: id, offset: 8 },
          });
          await pair.settle();
          for (const editor of pair.editors) {
            expect(editor.getReviewItems()).toHaveLength(1);
            expect(editor.getReviewItems({ pairReplacements: false })).toHaveLength(2);
            expect(editor.getReviewItems()[0]).toMatchObject({
              revisionKind: 'replace',
              text: 'Newer',
              replacedText: 'Old',
            });
          }
          pair.pause();
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: tail, offset: 4, text: ' Peer' },
            ]).rejected
          ).toBe(false);
          expect(alice![action](alice!.getReviewItems()[0]!.key).ok).toBe(true);
          pair.reconnect(first);
          await pair.settle();
          const expected = action === 'acceptReviewItem' ? 'Newer' : 'Old';
          for (const editor of pair.editors) {
            expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe(expected);
            expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Tail Peer');
            expect(editor.getReviewItems()).toHaveLength(0);
          }
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote).toBeDefined();
          expect(remote!.head.offset).toBe(bob!.surface!.state().selection!.head.offset);
          const position = { paragraphId: remote!.head.nodeId, offset: remote!.head.offset };
          const left = caretAt(alice!.surface!.layout(), position),
            right = caretAt(bob!.surface!.layout(), position);
          expect(left).not.toBeNull();
          expect(right).not.toBeNull();
          expect({ x: left!.x, y: left!.y }).toEqual({ x: right!.x, y: right!.y });
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('NewerOld');
            expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Tail Peer');
            expect(editor.getReviewItems()).toHaveLength(1);
          }
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe(expected);
            expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Tail Peer');
          }
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(await alice!.save()),
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          try {
            for (const editor of reopened.editors) {
              expect(
                paragraphTextOf(
                  editor.surface!.session.part(),
                  editor.surface!.session.paragraphIds()[0]!
                )
              ).toBe(expected);
              expect(editor.getReviewItems()).toHaveLength(0);
            }
          } finally {
            reopened.destroy();
          }
        } finally {
          pair.destroy();
        }
      }, 30_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    test(`leading replacement owner deletion preserves peer text on undo; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({
        document: leadingDocument(),
        clientIds,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors,
          id = alice!.surface!.session.paragraphIds()[0]!;
        pair.pause();
        expect(
          alice!.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]).rejected
        ).toBe(false);
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: id, offset: 8, text: ' Peer' },
          ]).rejected
        ).toBe(false);
        pair.reconnect(first);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBeNull();
          expect(editor.getReviewItems()).toHaveLength(0);
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('NewerOld Peer');
          expect(editor.getReviewItems()).toHaveLength(1);
        }
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBeNull();
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
      } finally {
        pair.destroy();
      }
    }, 30_000);
