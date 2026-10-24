// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { formattingFingerprint } from '../../test/support/formatting-fingerprint';
import { prepareNativeDocument } from '../native/prepare-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { insertManualList } from './insert-list';
import { setManualList } from './lists';

const context = { documentId: 'empty-list-peers', epoch: 'mount' };
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const [source, target] of [
    ['disc', 'disc'],
    ['decimal', 'decimal'],
    ['todo', 'todo'],
    ['disc', 'todo'],
    ['decimal', 'disc'],
    ['todo', 'decimal'],
  ] as const)
    for (const decision of ['direct', 'accept', 'reject'] as const)
      test(`empty ${source} to ${target} keeps offline peer text and complete Word shape: ${decision}/${clients}`, async () => {
        const seed = createDocxEditor({
          container: document.createElement('div'),
          document: docxFixture(
            clients[0] === 73 ? '<w:p><w:r><w:t>Task</w:t></w:r></w:p>' : '<w:p/>'
          ),
        });
        let bytes: Uint8Array;
        try {
          expect(setManualList(seed, source, context).success).toBe(true);
          bytes = new Uint8Array(await seed.save());
        } finally {
          seed.destroy();
        }
        const prepared = prepareNativeDocument(bytes);
        if (!prepared.success) throw new Error(prepared.error);
        const pair = await createTwoEditorFixture({
          document: prepared.document,
          documentId: context.documentId,
          clientIds: clients,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!;
          if (clients[0] === 73) {
            const paragraphId = a.session.paragraphIds()[0]!;
            a.setSelection({
              anchor: { paragraphId, offset: source === 'todo' ? 1 : 0 },
              head: { paragraphId, offset: (source === 'todo' ? 1 : 0) + 4 },
            });
            expect(a.deleteSelection()).toBe(true);
            await pair.settle();
          }
          const initial = a.session.part(),
            id = a.session.paragraphIds()[0]!;
          const point = { paragraphId: id, offset: source === 'todo' ? 1 : 0 };
          a.setSelection({ anchor: point, head: point });
          b.setSelection({ anchor: point, head: point });
          if (decision !== 'direct')
            expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          pair.pause();
          const inserted = insertManualList(alice!, target, context);
          if (!inserted.success) throw new Error(inserted.error);
          const sourceId = a.session.paragraphIds()[0]!;
          const expected = applyTreeOp(a.session.part(), {
            op: 'insertText',
            paragraphId: sourceId,
            offset: paragraphTextOf(a.session.part(), sourceId)!.length,
            text: 'Peer',
          });
          const restored = applyTreeOp(initial, {
            op: 'insertText',
            paragraphId: id,
            offset: point.offset,
            text: 'Peer',
          });
          if (!expected.ok || !restored.ok) throw new Error('Independent peer insertion refused.');
          b.insertPlainText('Peer');
          b.flushPendingInput();
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const after = canonicalOoxmlFingerprint(a.session.part());
          for (const peer of pair.editors)
            expect(formattingFingerprint(peer.surface!.session.part())).toEqual(
              formattingFingerprint(expected.part)
            );
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(formattingFingerprint(a.session.part())).toEqual(
            formattingFingerprint(restored.part)
          );
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(formattingFingerprint(a.session.part())).toEqual(formattingFingerprint(initial));
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(after);
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
          for (const peer of pair.editors) {
            const data = new Uint8Array(await peer.save());
            const read = readOoxmlPackage(data);
            if (!read.ok) throw new Error(read.reason);
            saved.push(semanticDigest(read.package.parts.values()));
            if (decision === 'reject')
              expect(
                formattingFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
              ).toEqual(formattingFingerprint(restored.part));
          }
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          pair.destroy();
        }
      });
