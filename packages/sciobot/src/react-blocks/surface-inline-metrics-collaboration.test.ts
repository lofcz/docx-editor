// The real-session transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { setSurfaceInlineWidgetMetrics } from '@docx-editor.dev/core/editor';
import { setPaginatedSurfaceScale } from '../../../core/src/editor/surface-scale';
import { caretAt, linesOf } from '@docx-editor.dev/core/layout';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';

const run = (text: string) => `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r>`;
const body = `<w:p>${run('A')}<w:sdt><w:sdtPr><w:tag w:val="chip:source"/></w:sdtPr><w:sdtContent>${run('source')}</w:sdtContent></w:sdt>${run('B')}</w:p><w:p>${run('Tail')}</w:p>`;
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`surface metric snapshots stay local through peer text, history, deletion, zoom, and save ${clients}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture(body),
        clientIds: clients,
      });
      try {
        const [alice, bob] = pair.editors,
          a = alice!.surface!,
          b = bob!.surface!;
        const id = a.session.paragraphIds()[0]!;
        const widget = (surface: typeof a) =>
          linesOf(surface.layout())
            .flatMap((line) => line.spans)
            .find((span) => span.inlineWidget)!;
        const before = pair.editors.map((editor) => ({
          revision: editor.surface!.session.packageRevision(),
          fingerprint: canonicalOoxmlFingerprint(editor.surface!.session.part().root),
        }));
        for (const [peer, surface] of [a, b].entries()) {
          setSurfaceInlineWidgetMetrics(surface, {
            measure: (request) =>
              request.tag === 'chip:source'
                ? { width: peer === 0 ? 64 : 40, ascent: 20, descent: 4 }
                : null,
            revision: '1',
          });
          expect(widget(surface).box.width).toBe(peer === 0 ? 48 : 30);
          expect(surface.session.packageRevision()).toBe(before[peer]!.revision);
          expect(canonicalOoxmlFingerprint(surface.session.part().root)).toBe(
            before[peer]!.fingerprint
          );
        }
        pair.pause();
        const start = { paragraphId: id, offset: 0 },
          end = { paragraphId: id, offset: 8 };
        a.setSelection({ anchor: start, head: start });
        a.insertPlainText('aa');
        a.flushPendingInput();
        b.setSelection({ anchor: end, head: end });
        b.insertPlainText('!');
        b.flushPendingInput();
        pair.reconnect(receiver);
        await pair.settle();
        for (const [peer, surface] of [a, b].entries()) {
          expect(paragraphTextOf(surface.session.part(), id)).toBe('aaAsourceB!');
          expect(widget(surface).range).toMatchObject({ start: 3, end: 9 });
          expect(caretAt(surface.layout(), { paragraphId: id, offset: 9 })?.x).toBe(
            peer === 0 ? 66 : 48
          );
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(id);
        expect(caretAt(a.layout(), { paragraphId: id, offset: remote!.head.offset })?.x).toBe(78);
        expect(setPaginatedSurfaceScale(a, 2)).toBe(true);
        expect(widget(a).box.width).toBe(48);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(paragraphTextOf(a.session.part(), id)).toBe('AsourceB!');
        expect(widget(a).box.width).toBe(48);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        pair.pause();
        const last = { paragraphId: id, offset: 11 };
        b.setSelection({ anchor: last, head: last });
        b.insertPlainText('?');
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
        pair.reconnect(receiver);
        await pair.settle();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const [peer, surface] of [a, b].entries()) {
          expect(paragraphTextOf(surface.session.part(), id)).toBe('aaAsourceB!?');
          expect(widget(surface).box.width).toBe(peer === 0 ? 48 : 30);
        }
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        setSurfaceInlineWidgetMetrics(a, null);
        expect(
          linesOf(a.layout())
            .flatMap((line) => line.spans)
            .some((span) => span.inlineWidget)
        ).toBe(false);
        expect(widget(b).box.width).toBe(30);
        expect((await pair.savedDigests())[0]).toEqual(saved[0]);
      } finally {
        pair.destroy();
      }
    });
