// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { newParagraphSeeds } from '../review/new-paragraphs';

const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const tracked of [false, true])
    test(`independent paragraph ${clients}/${tracked} preserves peer text across birth history and late redo`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'),
        clientIds: clients,
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const seeds = newParagraphSeeds([
          { type: 'paragraph', children: [{ type: 'text', value: 'Fresh' }] },
        ]);
        if (!seeds.success) throw Error(seeds.error);
        expect(
          alice!.surface!.applyAutomationOps(() => [
            {
              op: 'insertParagraphs',
              parentId: bodyStoryRoot(alice!.surface!.session.part())!.id,
              beforeId: null,
              paragraphs: seeds.paragraphs,
              ...(tracked
                ? { revision: { author: 'Alice', date: '2026-10-24T12:00:00.000Z' } }
                : {}),
            },
          ]).rejected
        ).toBe(false);
        await pair.settle();
        const born = bob!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Fresh')!;
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: born, offset: 2, text: ' peer ' },
          ]).rejected
        ).toBe(false);
        await pair.settle();
        const pending = pair.editors.map((editor) =>
          native(editor.surface!.session.currentPackage())
        );
        const check = (expected: string[]) =>
          pair.editors.forEach((editor) => {
            const session = editor.surface!.session;
            expect(session.paragraphIds().map((id) => paragraphTextOf(session.part(), id))).toEqual(
              expected
            );
          });
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(['Start', ' peer ']);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        pair.editors.forEach((editor, index) =>
          expect(native(editor.surface!.session.currentPackage())).toBe(pending[index]!)
        );
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(['Start', 'Fresh']);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(['Start']);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(['Start', ' peer ']);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await alice!.save()),
          ownedAttributes: true,
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
        });
        try {
          await reopened.settle();
          for (const editor of reopened.editors) {
            const session = editor.surface!.session;
            expect(session.paragraphIds().map((id) => paragraphTextOf(session.part(), id))).toEqual(
              ['Start', ' peer ']
            );
          }
        } finally {
          reopened.destroy();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    });
