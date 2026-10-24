// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt, selectionRects, spansInSelection } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { setTextColor, textColorState } from './colors';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const channel of ['foreground', 'background'] as const)
    for (const clear of [false, true])
      test(`${channel} ${clear ? 'clear' : 'RGB'} keeps peer text, history, geometry, reconnect and saved themes ${clients}`, async () => {
        const pair = await createTwoEditorFixture({
          clientIds: clients,
          document: docxFixture(
            '<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/><w:shd w:val="clear" w:fill="DDEEFF" w:themeFill="accent2"/></w:rPr><w:t>Imported</w:t></w:r></w:p><w:p><w:r><w:rPr><w:color w:val="CC0000"/><w:shd w:val="clear" w:fill="FFFF00"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p><w:p/>'
          ),
        });
        try {
          const [alice, bob] = pair.editors,
            a = alice!.surface!,
            b = bob!.surface!;
          const ids = a.session.paragraphIds(),
            id = ids[1]!;
          const selection = {
            anchor: { paragraphId: id, offset: 0 },
            head: { paragraphId: id, offset: 5 },
          };
          const originalSpan = () =>
            spansInSelection(a.layout(), selection, a.session.paragraphIds())[0]!;
          a.setSelection(selection);
          const beforeProps = originalSpan().props;
          b.setSelection({
            anchor: { paragraphId: id, offset: 10 },
            head: { paragraphId: id, offset: 10 },
          });
          pair.pause();
          expect(setTextColor(alice!, channel, clear ? null : '#17A2B8').ok).toBe(true);
          b.insertPlainText(' peer');
          b.flushPendingInput();
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const property = channel === 'foreground' ? 'color' : 'shd';
          const assertColor = () => {
            const current = originalSpan().props.find((p) => p.localName === property);
            if (clear) expect(current).toBeUndefined();
            else
              expect(current?.attributes?.[channel === 'foreground' ? 'val' : 'fill']).toBe(
                '17A2B8'
              );
          };
          assertColor();
          for (const editor of pair.editors) {
            const s = editor.surface!;
            s.setSelection(selection);
            expect(paragraphTextOf(s.session.part(), id)).toBe('Alpha beta peer');
            if (!clear) expect(textColorState(editor, channel).value).toBe('#17A2B8');
            expect(caretAt(s.layout(), { paragraphId: id, offset: 12 })).not.toBeNull();
            expect(
              selectionRects(s.layout(), selection, s.session.paragraphIds()).length
            ).toBeGreaterThan(0);
          }
          await pair.settle();
          expect(
            pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.nodeId
          ).toBe(id);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(originalSpan().props).toEqual(beforeProps);
          expect(paragraphTextOf(a.session.part(), id)).toBe('Alpha beta peer');
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          assertColor();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(paragraphTextOf(a.session.part(), id)).toBe('Alpha beta');
          assertColor();
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
            const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!saved.ok) throw new Error(saved.reason);
            expect(
              canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
            ).toEqual(canonicalOoxmlFingerprint(editor.surface!.session.part()));
          }
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          pair.destroy();
        }
      });
