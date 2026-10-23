// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { planParagraphText } from '../diff/paragraph-text';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { replaceOwnedProposals } from './replace-owned-proposals';
import { rejectReviewGroup } from './reject-group';
import { readReviewGroups } from './groups';
import { readWorkingView } from './working-view';
const word = (pkg: Parameters<typeof writeOoxmlPackage>[0]) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
};
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const concurrent of [false, true])
    test(`independent source scopes retain concurrent proposals, peer text, decisions, histories, and reopening; ${clients}/${concurrent}`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture(
          '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:r><w:t>Beta</w:t></w:r></w:p>'
        ),
        documentId: propertyContext.documentId,
        ownedAttributes: true,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const baseline = alice!.surface!.session.currentPackage(),
          id = alice!.surface!.session.paragraphIds()[0]!;
        pair.pause();
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: id, offset: 5, text: ' Peer' },
          ]).rejected
        ).toBe(false);
        const writes = pair.editors.map((editor, index) => {
          const text = paragraphTextOf(editor.surface!.session.part(), id)!;
          const plan = planParagraphText(
            {
              revision: propertyToken(editor),
              partName: baseline.mainDocumentPart,
              paragraphId: id,
              text,
            },
            index ? text.replace('Beta', 'Second') : text.replace('Alpha', 'First')
          );
          if (!plan.success) throw Error(plan.error);
          const result = replaceOwnedProposals(editor, [], plan.plan, propertyContext);
          if (!result.success) throw Error(result.error);
          return result;
        });
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const check = async () => {
          for (const editor of pair.editors) {
            const read = await readWorkingView(editor.surface!.session, propertyToken(editor));
            if (!read.success) throw Error(read.error);
            expect(read.view.projection.text.trimEnd()).toBe('First Peer Second');
          }
        };
        await check();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          const read = await readWorkingView(editor.surface!.session, propertyToken(editor));
          expect(read.success && read.view.projection.text.trimEnd()).toBe('Alpha Peer Second');
        }
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check();
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          const read = await readWorkingView(editor.surface!.session, propertyToken(editor));
          expect(read.success && read.view.projection.text.trimEnd()).toBe('First Peer Beta');
        }
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check();
        const peer = applyTreeOp(baseline.parts.get(baseline.mainDocumentPart)!, {
          op: 'insertText',
          paragraphId: id,
          offset: 5,
          text: ' Peer',
        });
        if (!peer.ok) throw Error(peer.reason);
        const expected = word(withPart(baseline, peer.part));
        const pending = pair.editors.map((editor) =>
          word(editor.surface!.session.currentPackage())
        );
        if (concurrent) pair.pause();
        for (const [index, write] of writes.entries()) {
          const editor = pair.editors[index]!;
          for (const group of write.ownership) {
            const result = rejectReviewGroup(
              editor,
              group.groupId,
              propertyToken(editor),
              propertyContext
            );
            if (!result.success) throw Error(result.error);
          }
          if (!concurrent) await pair.settle();
        }
        if (concurrent) pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        // Export a cold tree inside the observed transaction. Cached exports cannot mask writes.
        const roomBytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
        let exported: Uint8Array | undefined;
        const serialized = alice!.surface!.applyAutomationOps(
          () => [],
          { kind: 'body' },
          [
            (working) => {
              exported = writeOoxmlPackage(structuredClone(working));
              return working;
            },
          ],
          true,
          true,
          'after'
        );
        expect(serialized.rejected).toBe(false);
        await pair.settle();
        pair.documents.forEach((doc, index) =>
          expect(Y.encodeStateAsUpdate(doc)).toEqual(roomBytes[index]!)
        );
        const exportedRead = readOoxmlPackage(exported!);
        if (!exportedRead.ok) throw Error(exportedRead.reason);
        expect(
          canonicalOoxmlFingerprint(exportedRead.package.parts.get(baseline.mainDocumentPart)!)
        ).toBe(expected);
        for (const editor of pair.editors) {
          expect(word(editor.surface!.session.currentPackage())).toBe(expected);
          const rows = readReviewGroups(editor.surface!.session.currentPackage());
          expect(rows.success && rows.groups.length).toBe(0);
        }
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(word(editor.surface!.session.currentPackage())).toBe(expected);
        expect(pending[0]).toBe(pending[1]);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        const room = await createTwoEditorFixture({
          document: new Uint8Array(await bob!.save()),
          documentId: propertyContext.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
        });
        try {
          await room.settle();
          for (const editor of room.editors)
            expect(word(editor.surface!.session.currentPackage())).toBe(expected);
          room.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          room.destroy();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30000);
