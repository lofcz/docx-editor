import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { unzipSync, zipSync } from 'fflate';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  groupParagraph,
  groups,
  groupToken,
  proposeGroupText,
} from '../../test/support/review-groups';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroupStore } from '../review/group-record-store';
import { applyInlineCode, inlineCodeStyleState } from './inline-code';
import { prepareNativeDocument } from './prepare-document';

function insert(editor: DocxEditorInstance, paragraphId: string, text: string) {
  const surface = editor.surface!;
  surface.setSelection({ anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 0 } });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`missing owner relationships initialize before shared code/review edits, history, deletion, geometry, and reopen; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
      const entries = unzipSync(
        docxFixture(
          '<w:p><w:r><w:t>alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>The cat sat.</w:t></w:r></w:p>'
        )
      );
      delete entries['word/_rels/document.xml.rels'];
      const prepared = prepareNativeDocument(zipSync(entries));
      if (!prepared.success) throw new Error(prepared.error);
      const pair = await createTwoEditorFixture({
        document: prepared.document,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          codeId = groupParagraph(alice, 'alpha beta'),
          textId = groupParagraph(alice, 'The cat sat.');
        pair.pause();
        const coded = applyInlineCode(
          alice,
          { paragraphId: codeId, start: 0, end: 5, code: true },
          {
            documentId: propertyContext.documentId,
            epoch: propertyContext.epoch,
            revision: groupToken(alice),
          }
        );
        if (!coded.success) throw new Error(coded.error);
        proposeGroupText(bob, 'The cat sat.', 'The car sat.');
        const groupId = groups(bob)[0]!.id;
        bob.surface!.setSelection({
          anchor: { paragraphId: textId, offset: 2 },
          head: { paragraphId: textId, offset: 2 },
        });
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(inlineCodeStyleState(editor.surface!.session.currentPackage())).toBe('present');
          expect(readReviewGroupStore(editor.surface!.session.currentPackage()).status).toBe(
            'ready'
          );
          expect(groups(editor).map((group) => group.id)).toEqual([groupId]);
          expect(
            caretAt(editor.surface!.layout(), { paragraphId: textId, offset: 2 })
          ).not.toBeNull();
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(textId);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(inlineCodeStyleState(editor.surface!.session.currentPackage())).toBe('present');
          expect(groups(editor).map((group) => group.id)).toEqual([groupId]);
        }
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        pair.pause();
        insert(bob, codeId, 'Peer: ');
        const removed = alice.surface!.applyAutomationOps(
          () => [{ op: 'deleteBlock', blockId: codeId }],
          { kind: 'body' },
          undefined,
          false,
          true
        );
        expect(removed.rejected).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), codeId)).toBe('Peer: alpha beta');
          expect(groups(editor).map((group) => group.id)).toEqual([groupId]);
          const bytes = new Uint8Array(await editor.save()),
            repeated = prepareNativeDocument(bytes);
          if (!repeated.success) throw new Error(repeated.error);
          expect(repeated.changed).toBe(false);
          expect(repeated.document).toBe(bytes);
          const loaded = readOoxmlPackage(bytes);
          if (!loaded.ok) throw new Error(loaded.reason);
          const reopened = propertyEditor(loaded.package);
          try {
            expect(groups(reopened).map((group) => group.id)).toEqual([groupId]);
            expect(inlineCodeStyleState(reopened.surface!.session.currentPackage())).toBe(
              'present'
            );
          } finally {
            reopened.destroy();
          }
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    }, 60000);
