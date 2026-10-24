// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  applyTreeOp,
  paragraphTextOf,
  paraIdOf,
  findNode,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import { formattingFingerprint } from '../../test/support/formatting-fingerprint';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { rejectReviewGroup } from '../review/reject-group';
import { acceptReviewGroup } from '../review/accept-group';
import { readEditorRevision } from '../projection/revision-cache';
import { prepareNativeDocument } from '../native/prepare-document';
import { insertManualList } from './insert-list';
const context = { documentId: 'insert-lists', epoch: 'mount' };
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const style of ['disc', 'decimal', 'todo'] as const)
    for (const decision of ['direct', 'accept', 'reject'] as const)
      test(`list insertion ${style}/${decision} preserves offline text, independent undo, caret redo and saved convergence ${clients}`, async () => {
        const prepared = prepareNativeDocument(
          docxFixture(
            '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
          )
        );
        if (!prepared.success) throw new Error(prepared.error);
        const pair = await createTwoEditorFixture({
          clientIds: clients,
          documentId: context.documentId,
          ownedAttributes: true,
          document: prepared.document,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!,
            id = a.session.paragraphIds()[0]!;
          a.setSelection({
            anchor: { paragraphId: id, offset: 0 },
            head: { paragraphId: id, offset: 5 },
          });
          const point = { paragraphId: id, offset: 8 };
          b.setSelection({ anchor: point, head: point });
          if (decision !== 'direct')
            expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          const before = canonicalOoxmlFingerprint(a.session.part());
          pair.pause();
          const result = insertManualList(alice!, style, context);
          if (!result.success) throw new Error(result.error);
          const authored = a.session.part(),
            tail = a.session.paragraphIds()[1]!;
          const expectedPeer = applyTreeOp(authored, {
            op: 'insertText',
            paragraphId: tail,
            offset: 3,
            text: '!',
          });
          if (!expectedPeer.ok) throw new Error(expectedPeer.reason);
          b.insertPlainText('!');
          b.flushPendingInput();
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const after = canonicalOoxmlFingerprint(a.session.part()),
            selection = a.state().selection;
          expect(canonicalOoxmlFingerprint(b.session.part())).toEqual(after);
          expect(formattingFingerprint(a.session.part())).toEqual(
            formattingFingerprint(expectedPeer.part)
          );
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(a.session.paragraphIds().map((p) => paragraphTextOf(a.session.part(), p))).toEqual(
            ['Alpha be!ta', 'Tail']
          );
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(before);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(after);
          expect(a.state().selection).toEqual(selection);
          if (decision !== 'direct') {
            const groups = readReviewGroups(a.session.currentPackage());
            if (!groups.success) throw new Error(groups.error);
            expect(groups.groups).toHaveLength(1);
            const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
            expect(
              decide(
                alice!,
                groups.groups[0]!.id,
                readEditorRevision(alice!, context.documentId, context.epoch)!,
                context
              ).success
            ).toBe(true);
            await pair.settle();
          }
          const saved = [];
          for (const editor of pair.editors) {
            const read = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!read.ok) throw new Error(read.reason);
            saved.push(semanticDigest(read.package.parts.values()));
            const paragraphs: string[] = [];
            // Reopen through the native editor so retired shared paragraph addresses resolve as Word does.
            const fresh = createDocxEditor({
              container: document.createElement('div'),
              document: new Uint8Array(await editor.save()),
            });
            try {
              paragraphs.push(
                ...fresh
                  .surface!.session.paragraphIds()
                  .map((p) => paragraphTextOf(fresh.surface!.session.part(), p)!)
              );
              expect(paragraphs).toEqual(
                decision === 'reject'
                  ? ['Alpha be!ta', 'Tail']
                  : [style === 'todo' ? '☐' : '', ' be!ta', 'Tail']
              );
            } finally {
              fresh.destroy();
            }
          }
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          pair.destroy();
        }
      });

for (const style of ['disc', 'decimal'] as const)
  test(`shared ${style} insertion requires the production pre-room preparation without partial writes`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture('<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>'),
      documentId: context.documentId,
      ownedAttributes: true,
    });
    try {
      const editor = pair.editors[0]!,
        s = editor.surface!,
        before = canonicalOoxmlFingerprint(s.session.part()),
        revision = s.session.packageRevision(),
        selection = s.state().selection;
      expect(insertManualList(editor, style, context)).toMatchObject({
        success: false,
        error: 'This document is not ready for shared list insertion.',
      });
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(s.session.packageRevision()).toBe(revision);
      expect(s.state().selection).toEqual(selection);
      expect(editor.exec({ type: 'undo' }).ok).toBe(false);
      await pair.settle();
      expect(canonicalOoxmlFingerprint(pair.editors[1]!.surface!.session.part())).toEqual(before);
    } finally {
      pair.destroy();
    }
  });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const style of ['decimal', 'todo'] as const)
    test(`deleting an inserted ${style} item preserves concurrent typing, recovery, presence and Word output ${clients}`, async () => {
      const prepared = prepareNativeDocument(
        docxFixture(
          '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
        )
      );
      if (!prepared.success) throw new Error(prepared.error);
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        documentId: context.documentId,
        document: prepared.document,
        ownedAttributes: true,
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!,
          source = a.session.paragraphIds()[0]!;
        a.setSelection({
          anchor: { paragraphId: source, offset: 0 },
          head: { paragraphId: source, offset: 5 },
        });
        const inserted = insertManualList(alice!, style, context);
        if (!inserted.success) throw new Error(inserted.error);
        a.insertPlainText('Task');
        a.flushPendingInput();
        await pair.settle();
        const marker = style === 'todo' ? 1 : 0,
          id = a.session
            .paragraphIds()
            .find((p) => paraIdOf(findNode(a.session.part(), p)!) === inserted.paragraphIdentity)!,
          point = { paragraphId: id, offset: marker + 2 },
          before = a.session.part();
        a.setSelection({ anchor: point, head: point });
        b.setSelection({ anchor: point, head: point });
        pair.pause();
        const removed = a.applyAutomationOps(
          () => [{ op: 'deleteBlock', blockId: id }],
          { kind: 'body' },
          undefined,
          false,
          true
        );
        expect(removed.committed).toBe(true);
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const expected = applyTreeOp(before, {
          op: 'insertText',
          paragraphId: id,
          offset: point.offset,
          text: '!',
        });
        if (!expected.ok) throw new Error(expected.reason);
        const fingerprint = canonicalOoxmlFingerprint(expected.part);
        const selection = {
          anchor: { paragraphId: id, offset: marker },
          head: { paragraphId: id, offset: marker + 4 },
        };
        b.setSelection(selection);
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((p) => p.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        for (const editor of pair.editors) {
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(fingerprint);
          expect(caretAt(editor.surface!.layout(), selection.head)).not.toBeNull();
          expect(
            selectionRects(
              editor.surface!.layout(),
              selection,
              editor.surface!.session.paragraphIds()
            ).length
          ).toBeGreaterThan(0);
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          expect(
            canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
          ).toEqual(fingerprint);
        }
      } finally {
        pair.destroy();
      }
    });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const style of ['decimal', 'todo'] as const)
    test(`middle ${style} insertion keeps caret and Word identity after shared remapping ${clients}`, async () => {
      const prepared = prepareNativeDocument(
        docxFixture(
          '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
        )
      );
      if (!prepared.success) throw new Error(prepared.error);
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: prepared.document,
        documentId: context.documentId,
        ownedAttributes: true,
      });
      try {
        const editor = pair.editors[0]!,
          s = editor.surface!,
          id = s.session.paragraphIds()[0]!,
          point = { paragraphId: id, offset: 3 };
        s.setSelection({ anchor: point, head: point });
        const result = insertManualList(editor, style, context);
        if (!result.success) throw new Error(result.error);
        await pair.settle();
        const current = s.state().selection.head;
        expect(paraIdOf(findNode(s.session.part(), current.paragraphId)!)).toBe(
          result.paragraphIdentity
        );
        expect(current.offset).toBe(style === 'todo' ? 1 : 0);
        s.insertPlainText('Item');
        s.flushPendingInput();
        await pair.settle();
        for (const peer of pair.editors)
          expect(
            peer
              .surface!.session.paragraphIds()
              .map((p) => paragraphTextOf(peer.surface!.session.part(), p))
          ).toEqual(['Alp', style === 'todo' ? '☐Item' : 'Item', 'ha beta', 'Tail']);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });
