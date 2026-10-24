// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { setTextColor } from './colors';
import { formattingFingerprint } from '../../test/support/formatting-fingerprint';

for (const channel of ['foreground', 'background'] as const)
  for (const action of ['accept', 'reject'] as const)
    test(`${channel} tracked native formatting ${action} preserves peer text and independent native formatting semantics`, async () => {
      const bytes = docxFixture(
        '<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/><w:shd w:val="clear" w:fill="DDEEFF" w:themeFill="accent2"/><w:lang w:val="cs-CZ"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p>'
      );
      const pair = await createTwoEditorFixture({
        document: bytes,
        clientIds: [73, 41],
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      const reference = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!,
          r = reference.surface!;
        const id = a.session.paragraphIds()[0]!,
          rid = r.session.paragraphIds()[0]!;
        const selection = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 5 },
        };
        a.setSelection(selection);
        a.setEditingMode('suggest');
        r.setSelection({
          anchor: { paragraphId: rid, offset: 10 },
          head: { paragraphId: rid, offset: 10 },
        });
        r.insertPlainText(' peer');
        r.flushPendingInput();
        if (action === 'accept') {
          r.setSelection({
            anchor: { paragraphId: rid, offset: 0 },
            head: { paragraphId: rid, offset: 5 },
          });
          expect(setTextColor(reference, channel, '#17A2B8').ok).toBe(true);
        }
        const expected = formattingFingerprint(r.session.part());
        pair.pause();
        expect(setTextColor(alice!, channel, '#17A2B8').ok).toBe(true);
        b.setSelection({
          anchor: { paragraphId: id, offset: 10 },
          head: { paragraphId: id, offset: 10 },
        });
        b.insertPlainText(' peer');
        b.flushPendingInput();
        pair.reconnect(1);
        await pair.settle();
        for (const editor of pair.editors)
          expect(editor.getReviewItems().length).toBeGreaterThan(0);
        expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(editor.getReviewItems()).toHaveLength(0);
          expect(formattingFingerprint(editor.surface!.session.part())).toEqual(expected);
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          const savedPart = saved.package.parts.get(saved.package.mainDocumentPart)!;
          expect(formattingFingerprint(savedPart)).toEqual(expected);
          expect(canonicalOoxmlFingerprint(savedPart)).toEqual(
            canonicalOoxmlFingerprint(editor.surface!.session.part())
          );
        }
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(editor.getReviewItems().length).toBeGreaterThan(0);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(formattingFingerprint(editor.surface!.session.part())).toEqual(expected);
      } finally {
        reference.destroy();
        pair.destroy();
      }
    });
