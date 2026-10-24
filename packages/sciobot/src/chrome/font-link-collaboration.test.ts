// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { fontLinkDocument, selectFontLink } from '../../test/support/font-link-document';
import { setManualFontFamily, setManualFontSize, manualFontState } from './fonts';
import { applyManualLink, removeManualLink, createLinkToolbar } from './links';
import { savedTableFingerprint } from '../../test/support/manual-table-controls';
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const feature of ['fonts', 'font-reset', 'links'] as const)
    test(`${feature} retain peer text, native geometry, independent history, reconnect, deletion and exact Word output ${clients}`, async () => {
      const pair = await createTwoEditorFixture({
        clientIds: clients,
        document: fontLinkDocument(),
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        selectFontLink(alice!, 0, 5);
        if (feature === 'font-reset') {
          expect(setManualFontFamily(alice!, 'Arial').ok).toBe(true);
          expect(setManualFontSize(alice!, 9).ok).toBe(true);
          await pair.settle();
        }
        const before = a.session.part(),
          id = a.session.paragraphIds()[0]!;
        const peer = applyTreeOp(before, {
          op: 'insertText',
          paragraphId: id,
          offset: 2,
          text: '!',
        });
        if (!peer.ok) throw new Error(peer.reason);
        pair.pause();
        if (feature !== 'links') {
          expect(setManualFontFamily(alice!, feature === 'font-reset' ? null : 'Arial').ok).toBe(
            true
          );
          expect(setManualFontSize(alice!, feature === 'font-reset' ? null : 9).ok).toBe(true);
        } else expect(applyManualLink(alice!, 'https://example.com/old').ok).toBe(true);
        const authored = a.session.part();
        selectFontLink(bob!, 2);
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const after = applyTreeOp(authored, {
          op: 'insertText',
          paragraphId: id,
          offset: 2,
          text: '!',
        });
        if (!after.ok) throw new Error(after.reason);
        const expected = canonicalOoxmlFingerprint(after.part);
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(expected);
        if (feature === 'links') {
          selectFontLink(alice!, 2);
          expect(applyManualLink(alice!, 'https://example.com/new').ok).toBe(true);
          await pair.settle();
          const changed = canonicalOoxmlFingerprint(a.session.part());
          expect(a.hyperlinks.linkAtCaret()!.text).toBe('Al!pha');
          expect(b.hyperlinks.linksInCaretParagraph()[0]!.href).toBe('https://example.com/new');
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(expected);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(changed);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(removeManualLink(alice!).ok).toBe(true);
          await pair.settle();
          expect(paragraphTextOf(a.session.part(), id)).toBe('Al!pha beta');
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(expected);
        }
        const selection = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 6 },
        };
        b.setSelection(selection);
        await pair.settle();
        expect(
          pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
        ).toBe(id);
        for (const e of pair.editors) {
          const s = e.surface!;
          expect(caretAt(s.layout(), { paragraphId: id, offset: 2 })).not.toBeNull();
          expect(
            selectionRects(s.layout(), selection, s.session.paragraphIds()).length
          ).toBeGreaterThan(0);
          expect(await savedTableFingerprint(e)).toEqual(expected);
        }
        for (let i = 0; i < (feature !== 'links' ? 2 : 1); i++) {
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
        }
        for (const e of pair.editors)
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(peer.part)
          );
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(
          canonicalOoxmlFingerprint(before)
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (let i = 0; i < (feature !== 'links' ? 2 : 1); i++) {
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
        }
        expect(canonicalOoxmlFingerprint(a.session.part())).toEqual(expected);
        if (feature !== 'links') {
          selectFontLink(alice!, 0, 6);
          expect(manualFontState(alice!)).toEqual(
            feature === 'font-reset'
              ? { family: 'Calibri', sizePixels: 22 / 1.5 }
              : { family: 'Arial', sizePixels: 9 }
          );
        }
        selectFontLink(alice!, 0, 6);
        pair.pause();
        a.deleteForward();
        selectFontLink(bob!, 3);
        b.insertPlainText('?');
        b.flushPendingInput();
        pair.reconnect(clients[0] === 41 ? 1 : 0);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const recovered = applyTreeOp(after.part, {
          op: 'insertText',
          paragraphId: id,
          offset: 3,
          text: '?',
        });
        if (!recovered.ok) throw new Error(recovered.reason);
        for (const e of pair.editors) {
          expect(canonicalOoxmlFingerprint(e.surface!.session.part())).toEqual(
            canonicalOoxmlFingerprint(recovered.part)
          );
          expect(await savedTableFingerprint(e)).toEqual(canonicalOoxmlFingerprint(recovered.part));
        }
      } finally {
        pair.destroy();
      }
    });
test('a link dialog keeps its native target while an online peer adds a suffix', async () => {
  const pair = await createTwoEditorFixture({ document: fontLinkDocument() });
  const [alice, bob] = pair.editors,
    toolbar = createLinkToolbar(alice!);
  try {
    selectFontLink(alice!, 0, 5);
    toolbar.open();
    selectFontLink(bob!, 10);
    bob!.surface!.insertPlainText(' peer');
    bob!.surface!.flushPendingInput();
    await pair.settle();
    const input = toolbar.element.querySelector('input')!;
    input.value = 'https://example.com/dialog';
    Array.from(toolbar.element.querySelectorAll('button'))
      .find((b) => b.textContent === 'Apply link')!
      .click();
    await pair.settle();
    for (const e of pair.editors) {
      expect(e.surface!.hyperlinks.linksInCaretParagraph()[0]!.text).toBe('Alpha');
      expect(
        paragraphTextOf(e.surface!.session.part(), e.surface!.session.paragraphIds()[0]!)
      ).toBe('Alpha beta peer');
    }
  } finally {
    toolbar.dispose();
    pair.destroy();
  }
});
