// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  semanticDigest,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { listDocument, listParagraph } from '../../test/support/list-document';
import { setManualList, adjustManualListLevel } from './lists';
const context = { documentId: 'manual-lists', epoch: 'mount' };
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const style of ['decimal', 'todo'] as const)
    test(`mixed ${style} indentation keeps offline peer text and independent history ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        documentId: context.documentId,
        ownedAttributes: true,
        document: listDocument(
          listParagraph('List') + listParagraph('Plain') + listParagraph('Tail')
        ),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!,
          [list, plain] = a.session.paragraphIds();
        const point = { paragraphId: list!, offset: 0 };
        a.setSelection({ anchor: point, head: point });
        expect(setManualList(alice!, style, context).success).toBe(true);
        await pair.settle();
        a.setSelection({
          anchor: { paragraphId: plain!, offset: 5 },
          head: { paragraphId: list!, offset: style === 'todo' ? 1 : 0 },
        });
        const peerPoint = { paragraphId: plain!, offset: 2 };
        b.setSelection({ anchor: peerPoint, head: peerPoint });
        const before = a.session.part();
        pair.pause();
        expect(adjustManualListLevel(alice!, 'increase', context).success).toBe(true);
        const authored = a.session.part();
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const withPeer = (part: typeof before) => {
          const result = applyTreeOp(part, {
            op: 'insertText',
            paragraphId: plain!,
            offset: 2,
            text: '!',
          });
          if (!result.ok) throw new Error(result.reason);
          return canonicalOoxmlFingerprint(result.part);
        };
        const after = withPeer(authored);
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(after);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(
            withPeer(before)
          );
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(before)
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toEqual(after);
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          expect(semanticDigest(saved.package.parts.values())).toEqual(
            semanticDigest(editor.surface!.session.currentPackage().parts.values())
          );
        }
      } finally {
        pair.destroy();
      }
    });
