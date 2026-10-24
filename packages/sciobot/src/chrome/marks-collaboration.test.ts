// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import {
  paragraphTextOf,
  readOoxmlPackage,
  canonicalOoxmlFingerprint,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from '../native/prepare-document';
import { BASIC_MARKS, basicMarkState, toggleBasicMark } from './marks';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const mark of BASIC_MARKS)
    test(`${mark} keeps peer text, history, geometry, reconnect and saved structure ${clients}`, async () => {
      const prepared = prepareNativeDocument(
        docxFixture(
          '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
        )
      );
      if (!prepared.success) throw new Error(prepared.error);
      const pair = await createTwoEditorFixture({
        document: prepared.document,
        clientIds: clients,
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const id = a.session.paragraphIds()[0]!;
        const selected = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 5 },
        };
        a.setSelection(selected);
        const before = canonicalOoxmlFingerprint(a.session.part());
        b.setSelection({
          anchor: { paragraphId: id, offset: 10 },
          head: { paragraphId: id, offset: 10 },
        });
        pair.pause();
        expect(toggleBasicMark(alice!, mark).ok).toBe(true);
        b.insertPlainText(' peer');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        for (const surface of [a, b]) {
          surface.setSelection(selected);
          expect(paragraphTextOf(surface.session.part(), id)).toBe('Alpha beta peer');
          expect(caretAt(surface.layout(), { paragraphId: id, offset: 12 })).not.toBeNull();
          expect(
            selectionRects(surface.layout(), selected, surface.session.paragraphIds()).length
          ).toBeGreaterThan(0);
        }
        expect(basicMarkState(alice!, mark)).toBe('on');
        expect(basicMarkState(bob!, mark)).toBe('on');
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(basicMarkState(alice!, mark)).toBe('off');
        expect(paragraphTextOf(a.session.part(), id)).toBe('Alpha beta peer');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(basicMarkState(alice!, mark)).toBe('on');
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(paragraphTextOf(a.session.part(), id)).toBe('Alpha beta');
        expect(basicMarkState(alice!, mark)).toBe('on');
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        pair.pause();
        b.setSelection({
          anchor: { paragraphId: id, offset: 15 },
          head: { paragraphId: id, offset: 15 },
        });
        b.insertPlainText('!');
        b.flushPendingInput();
        expect(
          a.applyAutomationOps(
            () => [{ op: 'deleteBlock', blockId: id }],
            { kind: 'body' },
            undefined,
            false,
            true
          ).rejected
        ).toBe(false);
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Alpha beta peer!');
          editor.surface!.setSelection(selected);
          expect(basicMarkState(editor, mark)).toBe('on');
          const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!reopened.ok) throw new Error(reopened.reason);
          expect(
            paragraphTextOf(reopened.package.parts.get(reopened.package.mainDocumentPart)!, id)
          ).toBe('Alpha beta peer!');
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(b.session.part())
        );
        expect(canonicalOoxmlFingerprint(a.session.part())).not.toEqual(before);
      } finally {
        pair.destroy();
      }
    });
