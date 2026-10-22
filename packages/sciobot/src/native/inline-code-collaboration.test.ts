import { expect, test } from 'bun:test';
import { encodeStateAsUpdate } from 'yjs';
import { strFromU8, unzipSync } from 'fflate';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import { findNode, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { readEditorRevision } from '../projection/revision-cache';
import { readCommittedView } from '../review/committed-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readWorkingView } from '../review/working-view';
import { applyInlineCode, type InlineCodeRequest } from './inline-code';
import { prepareNativeDocument } from './prepare-document';

const BODY =
  '<w:p><w:r><w:t xml:space="preserve">alpha beta gamma</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">tail text</w:t></w:r></w:p>';
const DOCUMENT = docxFixture(BODY);
const STYLED = docxFixture(BODY, {
  styles:
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '</w:styles>',
});
const IMPORTED = docxFixture(
  BODY.replaceAll(
    '<w:r>',
    '<w:r><w:rPr><w:rStyle w:val="Imported"/><w:customKeep w:val="keep"/></w:rPr>'
  ),
  {
    styles:
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:style w:type="character" w:styleId="Imported"><w:name w:val="Imported"/><w:rPr><w:b/></w:rPr></w:style></w:styles>',
  }
);
const prepared = (document: Uint8Array) => {
  const result = prepareNativeDocument(document);
  if (!result.success) throw new Error(result.error);
  return result.document;
};
const ids = { documentId: 'inline-code-room', epoch: 'mount-1' };
const token = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, ids.documentId, ids.epoch)!;
function code(editor: DocxEditorInstance, request: InlineCodeRequest) {
  const result = applyInlineCode(editor, request, { ...ids, revision: token(editor) });
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
}
async function views(editor: DocxEditorInstance) {
  const original = await readCommittedView(editor.surface!.session, token(editor));
  const working = await readWorkingView(editor.surface!.session, token(editor));
  if (!original.success || !working.success) throw new Error('Missing review view');
  return [original.view.projection.text, working.view.projection.text];
}
async function codeDefinitions(editor: DocxEditorInstance): Promise<number> {
  const entries = unzipSync(new Uint8Array(await editor.save()));
  const styles = entries['word/styles.xml'];
  return styles ? (strFromU8(styles).match(/w:styleId="SciobotCode"/g) ?? []).length : 0;
}

for (const clientIds of [
  [73, 41],
  [41, 73],
] as const)
  for (const receiver of [0, 1] as const)
    for (const actor of [0, 1] as const)
      test(`concurrent first code writes keep one shared styles part and independent history; clients ${clientIds.join('/')}, receiver ${receiver}, actor ${actor}`, async () => {
        // The application prepares original bytes before creating a shared room. The original has
        // no styles part; this same initializer also runs on ordinary native workspace imports.
        expect(unzipSync(DOCUMENT)['word/styles.xml']).toBeUndefined();
        const prepared = prepareNativeDocument(DOCUMENT);
        if (!prepared.success) throw new Error(prepared.error);
        const pair = await createTwoEditorFixture({
          document: prepared.document,
          clientIds,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors;
          const [first, second] = alice!.surface!.session.paragraphIds();
          for (const room of pair.rooms) expect(room.session.canUndo()).toBe(false);
          pair.pause();
          code(alice!, { paragraphId: first!, start: 6, end: 10, code: true });
          code(bob!, { paragraphId: second!, start: 0, end: 4, code: true });
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect((await views(editor))[1]).toBe('alpha `beta` gamma\n\n`tail` text');
            expect(await codeDefinitions(editor)).toBe(1);
          }
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          const owner = pair.editors[actor]!;
          const withoutOwner =
            actor === 0 ? 'alpha beta gamma\n\n`tail` text' : 'alpha `beta` gamma\n\ntail text';
          expect(owner.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect((await views(editor))[1]).toBe(withoutOwner);
            expect(await codeDefinitions(editor)).toBe(1);
          }
          for (const room of pair.rooms) expect(room.session.status()).toBe('ready');
          expect(owner.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect((await views(editor))[1]).toBe('alpha `beta` gamma\n\n`tail` text');
          // A new code write remains its own history step after the concurrent first writes.
          code(owner, {
            paragraphId: actor === 0 ? first! : second!,
            start: actor === 0 ? 6 : 0,
            end: actor === 0 ? 10 : 4,
            code: false,
          });
          await pair.settle();
          for (const editor of pair.editors) expect((await views(editor))[1]).toBe(withoutOwner);
          expect(owner.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
          const bytes = new Uint8Array(await alice!.save());
          for (const roomSnapshot of [undefined, encodeStateAsUpdate(pair.documents[0]!)]) {
            const reopened = await createTwoEditorFixture({
              document: bytes,
              roomSnapshot,
              ownedAttributes: true,
              editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            });
            try {
              for (const editor of reopened.editors) {
                expect((await views(editor))[1]).toBe('alpha `beta` gamma\n\n`tail` text');
                expect(await codeDefinitions(editor)).toBe(1);
              }
              for (const room of reopened.rooms) expect(room.session.status()).toBe('ready');
              const restored = await reopened.savedDigests();
              expect(restored[0]).toEqual(restored[1]);
            } finally {
              reopened.destroy();
            }
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`derived code styles retain imported marks, peer text, history, and geometry; clients ${clientIds.join('/')}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: prepared(IMPORTED),
        clientIds,
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const [first, second] = alice!.surface!.session.paragraphIds();
        pair.pause();
        code(alice!, { paragraphId: first!, start: 6, end: 10, code: true });
        code(bob!, { paragraphId: second!, start: 0, end: 4, code: true });
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: first!, offset: 16, text: ' Bob' },
          ]).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        const landed = '**alpha `beta` gamma Bob**\n\n**`tail` text**';
        for (const editor of pair.editors) {
          expect((await views(editor))[1]).toBe(landed);
          pair.rooms.forEach((room) => expect(room.session.status()).toBe('ready'));
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect((await views(editor))[1]).toBe('**alpha beta gamma Bob**\n\n**`tail` text**');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        bob!.surface!.setSelection({
          anchor: { paragraphId: first!, offset: 6 },
          head: { paragraphId: first!, offset: 10 },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (entry) => entry.actorId === 'bob'
        )!;
        expect(
          caretAt(alice!.surface!.layout(), {
            paragraphId: remote.head.nodeId,
            offset: remote.head.offset,
          })
        ).toEqual(caretAt(bob!.surface!.layout(), { paragraphId: first!, offset: 10 }));
        expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
        pair.pause();
        code(alice!, { paragraphId: second!, start: 0, end: 4, code: false });
        expect(
          bob!.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: second! }]).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect((await views(editor))[1]).toBe('**alpha `beta` gamma Bob**');
          expect(findNode(editor.surface!.session.part(), second!)).toBeNull();
        }
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 60_000);

for (const action of ['accept', 'reject'] as const)
  for (const receiver of [0, 1] as const)
    test(`derived code formatting preserves imported style through ${action}, history, and reopen; receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: prepared(IMPORTED),
        clientIds: [73, 41],
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const [first] = alice!.surface!.session.paragraphIds();
        alice!.surface!.setEditingMode('suggest');
        pair.pause();
        code(alice!, { paragraphId: first!, start: 6, end: 10, code: true });
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: first!, offset: 16, text: ' Bob' },
          ]).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        const original = '**alpha beta gamma Bob**\n\n**tail text**';
        const proposed = '**alpha `beta` gamma Bob**\n\n**tail text**';
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([original, proposed]);
        expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        await pair.settle();
        const resolved = action === 'accept' ? proposed : original;
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([resolved, resolved]);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([original, proposed]);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await alice!.save()),
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          for (const editor of reopened.editors)
            expect(await views(editor)).toEqual([resolved, resolved]);
        } finally {
          reopened.destroy();
        }
      } finally {
        pair.destroy();
      }
    }, 60_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const styled of [true, false])
      test(`concurrent code, peer text, history, geometry, and deletion converge; ${styled ? 'both peers add the style' : 'one peer creates the styles part'}; clients ${clientIds.join('/')}, receiver ${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: styled ? STYLED : DOCUMENT,
          documentId: `inline-code-${clientIds.join('-')}-${receiver}-${styled}`,
          clientIds,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors;
          const [first, second] = alice!.surface!.session.paragraphIds();
          pair.pause();
          // Both peers add the code style offline when the styles part exists. Concurrent creation
          // of a missing styles part is a transport part contest, so only Alice creates it there.
          // Bob always edits around Alice's range.
          code(alice!, { paragraphId: first!, start: 6, end: 10, code: true });
          if (styled) code(bob!, { paragraphId: second!, start: 0, end: 4, code: true });
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: first!, offset: 16, text: ' Bob' },
              { op: 'deleteText', paragraphId: first!, start: 0, end: 6 },
            ]).rejected
          ).toBe(false);
          pair.reconnect(receiver);
          await pair.settle();
          const tail = styled ? '`tail` text' : 'tail text';
          const landed = `\`beta\` gamma Bob\n\n${tail}`;
          for (const editor of pair.editors) {
            expect((await views(editor))[1]).toBe(landed);
            // Each peer's concurrent definition survives the merge until the next code write.
            expect(await codeDefinitions(editor)).toBe(styled ? 2 : 1);
          }

          // Alice's history removes only her code; Bob's text and code stay.
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect((await views(editor))[1]).toBe(`beta gamma Bob\n\n${tail}`);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect((await views(editor))[1]).toBe(landed);

          // Remote selection geometry over the code span matches in both sessions.
          bob!.surface!.setSelection({
            anchor: { paragraphId: first!, offset: 1 },
            head: { paragraphId: first!, offset: 4 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (entry) => entry.actorId === 'bob'
          )!;
          expect(remote.head).toMatchObject({ nodeId: first, offset: 4 });
          const local = caretAt(bob!.surface!.layout(), { paragraphId: first!, offset: 4 });
          expect(local).not.toBeNull();
          expect(
            caretAt(alice!.surface!.layout(), {
              paragraphId: remote.head.nodeId,
              offset: remote.head.offset,
            })
          ).toEqual(local);
          const range = {
            anchor: { paragraphId: first!, offset: 1 },
            head: { paragraphId: first!, offset: 4 },
          };
          const rects = selectionRects(
            alice!.surface!.layout(),
            range,
            alice!.surface!.session.paragraphIds()
          );
          expect(rects.length).toBeGreaterThan(0);
          expect(rects).toEqual(
            selectionRects(bob!.surface!.layout(), range, bob!.surface!.session.paragraphIds())
          );
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);

          // Deletion during peer editing: Alice changes code while Bob deletes the paragraph.
          pair.pause();
          code(alice!, { paragraphId: second!, start: 0, end: 4, code: !styled });
          expect(
            bob!.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: second! }])
              .rejected
          ).toBe(false);
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(findNode(editor.surface!.session.part(), second!)).toBeNull();
            expect((await views(editor))[1]).toBe('`beta` gamma Bob');
            // That write also kept exactly one code style definition.
            expect(await codeDefinitions(editor)).toBe(1);
          }
          const reconciled = await pair.savedDigests();
          expect(reconciled[0]).toEqual(reconciled[1]);
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);

for (const action of ['accept', 'reject'] as const)
  for (const receiver of [0, 1] as const)
    test(`suggested code is visible to the peer and its decision converges (${action}); receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: DOCUMENT,
        documentId: `inline-code-review-${action}-${receiver}`,
        clientIds: [73, 41],
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const [first] = alice!.surface!.session.paragraphIds();
        alice!.surface!.setEditingMode('suggest');
        pair.pause();
        code(alice!, { paragraphId: first!, start: 6, end: 10, code: true });
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: first!, offset: 16, text: ' Bob' },
          ]).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        const pending = [
          'alpha beta gamma Bob\n\ntail text',
          'alpha `beta` gamma Bob\n\ntail text',
        ];
        for (const editor of pair.editors) {
          expect(await views(editor)).toEqual(pending);
          expect(editor.getReviewItems()).toHaveLength(1);
        }
        expect(alice!.surface!.editingMode()).toBe('suggest');
        expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        await pair.settle();
        const resolved = action === 'accept' ? pending[1]! : pending[0]!;
        for (const editor of pair.editors) {
          expect(await views(editor)).toEqual([resolved, resolved]);
          expect(editor.getReviewItems()).toHaveLength(0);
          expect(paragraphTextOf(editor.surface!.session.part(), first!)).toBe(
            'alpha beta gamma Bob'
          );
        }
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(await views(editor)).toEqual(pending);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([resolved, resolved]);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await alice!.save()),
          documentId: `inline-code-reopen-${action}-${receiver}`,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          for (const editor of reopened.editors)
            expect(await views(editor)).toEqual([resolved, resolved]);
        } finally {
          reopened.destroy();
        }
      } finally {
        pair.destroy();
      }
    }, 60_000);

for (const clientIds of [
  [73, 41],
  [41, 73],
] as const)
  for (const receiver of [0, 1] as const)
    test(`prepared derived styles survive a later code write and peer undo; clients ${clientIds.join('/')}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: prepared(IMPORTED),
        clientIds,
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const [first, second] = alice!.surface!.session.paragraphIds();
        for (const room of pair.rooms) expect(room.session.canUndo()).toBe(false);
        pair.pause();
        code(alice!, { paragraphId: first!, start: 6, end: 10, code: true });
        code(bob!, { paragraphId: second!, start: 0, end: 4, code: true });
        pair.reconnect(receiver);
        await pair.settle();
        code(alice!, { paragraphId: first!, start: 0, end: 5, code: true });
        await pair.settle();
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect((await views(editor))[1]).toBe('**`alpha` `beta` gamma**\n\n**tail text**');
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        const landed = '**`alpha` `beta` gamma**\n\n**`tail` text**';
        for (const editor of pair.editors) expect((await views(editor))[1]).toBe(landed);
        for (const room of pair.rooms) expect(room.session.status()).toBe('ready');
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await alice!.save()),
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          for (const editor of reopened.editors) expect((await views(editor))[1]).toBe(landed);
          const saved = await reopened.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          reopened.destroy();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 60_000);
