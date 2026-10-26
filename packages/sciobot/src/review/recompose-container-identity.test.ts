// Shared transport is evaluated only in tests under the EigenPal Pro License.
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { buildBlockIndex } from '../blocks/block-index';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { readReviewGroups } from './groups';
import { readPropertyReviewStore } from './property-record-store';
import { decodeBlockPresenceRecord } from './block-presence-schema';
import { recomposeOwnedProposals } from './recompose';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';

const context = {
  documentId: 'container-identity',
  epoch: 'native',
  author: 'AI',
  date: '2026-10-26T12:00:00Z',
};
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const token = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, context.documentId, context.epoch)!;
function fixture(restriction = '') {
  const read = readOoxmlPackage(
    docxFixture(
      p('Lead') +
        `<w:sdt><w:sdtPr><w:tag w:val="${blockTag('excerpt')}"/>${restriction}</w:sdtPr><w:sdtContent>${p('First quoted sentence.')}${p('Second quoted sentence.')}</w:sdtContent></w:sdt>` +
        p('Tail')
    )
  );
  if (!read.ok) throw Error(read.reason);
  const stored = withPayloadRecord(
    read.package,
    {
      version: 1,
      id: 'excerpt',
      type: 'source_excerpt',
      attributes: {
        title: 'Polednice',
        author: 'Erben',
        bookId: 123,
        url: 'https://source.test/book',
        trust: 'verified',
        sourceType: 'web',
        opaque: 'keep',
      },
    },
    'Excerpt'
  );
  if (!stored.ok) throw Error(stored.reason);
  return writeOoxmlPackage(stored.pkg);
}
const mount = (bytes = fixture()) =>
  createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
async function view(editor: DocxEditorInstance, committed = false) {
  const read = await (committed ? readCommittedView : readWorkingView)(
    editor.surface!.session,
    token(editor)
  );
  if (!read.success) throw Error(read.error);
  return read.view.projection.text;
}
for (const deletion of [false, true])
  for (const decision of ['accept', 'reject'] as const)
    test(`retained container ${deletion ? 'removal' : 'identity replacement'} uses exact structural presence through ${decision} and history`, async () => {
      const editor = mount();
      let cold: DocxEditorInstance | undefined;
      try {
        const initial = await view(editor);
        const before = editor.surface!.session.currentPackage();
        const original = buildBlockIndex(before).byId.get('excerpt')!;
        const native = canonicalOoxmlFingerprint(
          findNode(editor.surface!.session.part(), original.controlId)!
        );
        const target = deletion ? 'Lead\n\nTail' : initial.replace('Polednice', 'Kytice');
        const result = await recomposeOwnedProposals(
          editor,
          { revision: token(editor), targetMarkdown: target, ownership: [], contentProofs: [] },
          context,
          { createPresenceProposalId: () => 'container-change' }
        );
        if (!result.success) throw Error(result.error);
        expect(await view(editor)).toBe(target);
        expect(await view(editor, true)).toBe(initial);
        expect(
          canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), original.controlId)!)
        ).toBe(native);
        const records = readPropertyReviewStore(editor.surface!.session.currentPackage());
        if (records.status !== 'ready') throw Error('Missing presence store');
        const decoded = decodeBlockPresenceRecord(records.records.get('container-change')!.raw);
        if (decoded.status !== 'supported') throw Error('Missing block presence');
        expect(decoded.value.schemaVersion).toBe(2);
        expect(decoded.value.blocks).toHaveLength(deletion ? 1 : 2);
        expect(decoded.value.blocks).toContainEqual(
          expect.objectContaining({
            blockId: 'excerpt',
            before: { present: true },
            after: { present: false },
          })
        );
        const groups = readReviewGroups(editor.surface!.session.currentPackage());
        if (!groups.success) throw Error(groups.error);
        expect(groups.groups.every((group) => !group.readOnly && !group.native.length)).toBe(true);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(await view(editor)).toBe(initial);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(await view(editor)).toBe(target);
        cold = mount(new Uint8Array(await editor.save()));
        expect(await view(cold)).toBe(target);
        expect(await view(cold, true)).toBe(initial);
        const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          editor,
          token(editor),
          context
        );
        if (!decided.success) throw Error(decided.error);
        const expected = decision === 'accept' ? target : initial;
        expect(await view(editor)).toBe(expected);
        expect(await view(editor, true)).toBe(expected);
        const payloads = [
          ...buildBlockIndex(editor.surface!.session.currentPackage()).byId.values(),
        ];
        expect(payloads).toHaveLength(decision === 'accept' && deletion ? 0 : 1);
        if (payloads.length)
          expect(payloads[0]!.payload.attributes).toEqual({
            ...original.payload.attributes,
            title: decision === 'accept' ? 'Kytice' : 'Polednice',
          });
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(await view(editor)).toBe(target);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(await view(editor)).toBe(expected);
        if (!deletion || decision === 'reject')
          expect(
            editor
              .surface!.session.paragraphIds()
              .map((id) => paragraphTextOf(editor.surface!.session.part(), id))
          ).toEqual(['Lead', 'First quoted sentence.', 'Second quoted sentence.', 'Tail']);
      } finally {
        cold?.destroy();
        editor.destroy();
      }
    });

for (const restriction of [
  '<w:lock w:val="contentLocked"/>',
  '<w:lock w:val="sdtLocked"/>',
  '<w:temporary/>',
  '<w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000000}" w:xpath="/quote"/>',
])
  test(`protected container replacement refuses before publication: ${restriction}`, async () => {
    const editor = mount(fixture(restriction));
    try {
      const initial = await view(editor),
        before = editor.surface!.session.currentPackage();
      let publications = 0;
      const stop = editor.surface!.session.subscribe(() => publications++);
      const result = await recomposeOwnedProposals(
        editor,
        {
          revision: token(editor),
          targetMarkdown: initial.replace('Polednice', 'Kytice'),
          ownership: [],
          contentProofs: [],
        },
        context,
        { createPresenceProposalId: () => 'protected-change' }
      );
      stop();
      expect(result.success).toBe(false);
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(publications).toBe(0);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(await view(editor)).toBe(initial);
    } finally {
      editor.destroy();
    }
  });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`container identity waves preserve peer text, histories, geometry and cold rooms: ${clients}/${receiver}/${decision}`, async () => {
        const pair = await createTwoEditorFixture({
          document: fixture(),
          documentId: context.documentId,
          ownedAttributes: true,
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        let cold: Awaited<ReturnType<typeof createTwoEditorFixture>> | undefined;
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const source = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get(
            'excerpt'
          )!;
          const bodyId = alice
            .surface!.session.paragraphIds()
            .find(
              (id) =>
                paragraphTextOf(alice.surface!.session.part(), id) === 'First quoted sentence.'
            )!;
          expect(
            alice.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId: bodyId, offset: 22, text: ' manual' }],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          await pair.settle();
          const initial = await view(alice);
          const original = canonicalOoxmlFingerprint(
            findNode(alice.surface!.session.part(), source.controlId)!
          );
          const firstTarget = initial.replace('Polednice', 'First book');
          const first = await recomposeOwnedProposals(
            alice,
            {
              revision: token(alice),
              targetMarkdown: firstTarget,
              ownership: [],
              contentProofs: [],
            },
            context,
            { createPresenceProposalId: () => 'first-container' }
          );
          if (!first.success) throw Error(first.error);
          await pair.settle();
          pair.pause();
          const tail = bob
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Tail')!;
          expect(
            bob.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: 'x' }],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          const target = firstTarget.replace('First book', 'Kytice');
          const second = await recomposeOwnedProposals(
            alice,
            {
              revision: token(alice),
              targetMarkdown: target,
              ownership: first.ownership,
              contentProofs: first.contentProofs,
            },
            context,
            { createPresenceProposalId: () => 'second-container' }
          );
          if (!second.success) throw Error(second.error);
          pair.reconnect(receiver);
          await pair.settle();
          const originalView = initial.replace('Tail', 'Tailx'),
            proposed = target.replace('Tail', 'Tailx');
          for (const editor of pair.editors) {
            expect(await view(editor)).toBe(proposed);
            expect(await view(editor, true)).toBe(originalView);
            const old = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(
              'excerpt'
            )!;
            expect(old.payload).toEqual(source.payload);
            expect(
              canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), old.controlId)!)
            ).toBe(original);
          }
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await view(bob)).toBe(firstTarget.replace('Tail', 'Tailx'));
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await view(bob)).toBe(proposed);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await view(alice)).toBe(target);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await view(alice)).toBe(proposed);
          const point = { paragraphId: bodyId, offset: 3 };
          bob.surface!.setSelection({ anchor: point, head: point });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (entry) => entry.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(bodyId);
          expect(remote?.head.offset).toBe(3);
          const a = caretAt(alice.surface!.layout(), point),
            b = caretAt(bob.surface!.layout(), point);
          expect(a).not.toBeNull();
          expect(b).not.toBeNull();
          expect({ x: a!.x, y: a!.y, height: a!.height }).toEqual({
            x: b!.x,
            y: b!.y,
            height: b!.height,
          });
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
          cold = await createTwoEditorFixture({
            document: fixture(),
            documentId: context.documentId,
            ownedAttributes: true,
            clientIds: [101, 303],
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          for (const editor of cold.editors) {
            expect(await view(editor)).toBe(proposed);
            expect(await view(editor, true)).toBe(originalView);
          }
          const result = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            alice,
            token(alice),
            context
          );
          if (!result.success) throw Error(result.error);
          await pair.settle();
          const expected = decision === 'accept' ? proposed : originalView;
          for (const editor of pair.editors) {
            expect(await view(editor)).toBe(expected);
            expect(await view(editor, true)).toBe(expected);
            const excerpts = [
              ...buildBlockIndex(editor.surface!.session.currentPackage()).byId.values(),
            ].filter((block) => block.payload.type === 'source_excerpt');
            expect(excerpts).toHaveLength(1);
            expect(excerpts[0]!.payload.attributes).toEqual({
              ...source.payload.attributes,
              title: decision === 'accept' ? 'Kytice' : 'Polednice',
            });
          }
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(await view(bob)).toBe(proposed);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(await view(bob)).toBe(expected);
          const final = await pair.savedDigests();
          expect(final[0]).toEqual(final[1]);
        } finally {
          cold?.destroy();
          pair.destroy();
        }
      });
