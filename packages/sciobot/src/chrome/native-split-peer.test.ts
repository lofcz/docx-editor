// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import { formattingFingerprint } from '../../test/support/formatting-fingerprint';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { captureReviewGroupHistory } from '../review/group-history';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { checklistSplitBudget } from './list-split-tracking';
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const collected of [false, true])
    for (const ownedAttributes of [false, true])
      for (const tracked of [false, true])
        test(`native delete/split keeps suffix placement through both actors' undo/redo ${clients}/${ownedAttributes}/${tracked}/gc=${collected}`, async () => {
          const pair = await createTwoEditorFixture({
            clientIds: clients,
            document: docxFixture('<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>'),
            ownedAttributes,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          pair.documents[0]!.gc = collected;
          pair.documents[1]!.gc = !collected;
          try {
            const [alice, bob] = pair.editors,
              a = alice!.surface!,
              b = bob!.surface!,
              id = a.session.paragraphIds()[0]!;
            const point = { paragraphId: id, offset: 5 },
              peer = { paragraphId: id, offset: 8 };
            a.setSelection({ anchor: point, head: point });
            b.setSelection({ anchor: peer, head: peer });
            if (tracked)
              expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
            const revision = {
              author: a.author() ?? '',
              date: `${new Date().toISOString().slice(0, 19)}Z`,
            };
            const removal = {
                op: 'deleteText' as const,
                paragraphId: id,
                start: 0,
                end: 5,
                ...(tracked ? { revision } : {}),
              },
              split = { op: 'splitParagraph' as const, paragraphId: id, offset: tracked ? 5 : 0 };
            const mark = {
              op: 'setParagraphMarkRevision' as const,
              paragraphId: id,
              kind: 'ins' as const,
              revision,
            };
            pair.pause();
            const execute = () =>
              a.applyAutomationOps(
                () => [removal],
                { kind: 'body' },
                undefined,
                tracked,
                true,
                'after',
                (ctx) => {
                  if (!ctx.apply(split) || (tracked && !ctx.apply(mark)))
                    throw new Error('The native split refused.');
                }
              );
            if (ownedAttributes) {
              const budget = checklistSplitBudget(
                a.session.currentPackage(),
                a.session.part().name,
                tracked ? [removal, split, mark] : [removal, split],
                (pkg) => pkg
              );
              expect(
                captureReviewGroupHistory(
                  a.collaborationSession(),
                  execute,
                  budget.flags,
                  budget.fields
                ).committed
              ).toBe(true);
            } else expect(execute().committed).toBe(true);
            b.insertPlainText('!');
            b.flushPendingInput();
            pair.reconnect(clients[0] === 41 ? 0 : 1);
            await pair.settle();
            const after = canonicalOoxmlFingerprint(a.session.part());
            expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            expect(
              a.session.paragraphIds().map((p) => paragraphTextOf(a.session.part(), p))
            ).toEqual(['Alpha be!ta']);
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(after);
            const saved = await pair.savedDigests();
            expect(saved[0]).toEqual(saved[1]);
          } finally {
            pair.destroy();
          }
        });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`concurrent deletion preserves the full Word shape of ordinary peer text without a split ${clients}`, async () => {
    const pair = await createTwoEditorFixture({
      clientIds: clients,
      document: docxFixture('<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>'),
      ownedAttributes: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!;
      const id = a.session.paragraphIds()[0]!;
      a.setSelection({
        anchor: { paragraphId: id, offset: 0 },
        head: { paragraphId: id, offset: 5 },
      });
      const point = { paragraphId: id, offset: 8 };
      b.setSelection({ anchor: point, head: point });
      expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      pair.pause();
      expect(a.deleteSelection()).toBe(true);
      const expected = applyTreeOp(a.session.part(), {
        op: 'insertText',
        paragraphId: id,
        offset: 8,
        text: '!',
      });
      if (!expected.ok) throw new Error(expected.reason);
      b.insertPlainText('!');
      b.flushPendingInput();
      pair.reconnect(clients[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors)
        expect(formattingFingerprint(editor.surface!.session.part())).toEqual(
          formattingFingerprint(expected.part)
        );
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
    } finally {
      pair.destroy();
    }
  });
