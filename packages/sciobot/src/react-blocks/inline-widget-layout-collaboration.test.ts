// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import {
  caretAt,
  createFixedMeasurer,
  createLayoutSession,
  createParagraphLayoutCache,
  layoutSemanticDocument,
  linesOf,
  type TextMeasurer,
} from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';

const run = (text: string) => `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r>`;
const body = `<w:p>${run('A')}<w:sdt><w:sdtPr><w:tag w:val="chip:source"/></w:sdtPr><w:sdtContent>${run('source')}</w:sdtContent></w:sdt>${run('B')}</w:p><w:p>${run('Tail')}</w:p>`;
const geometry = { width: 100, height: 200, margin: { top: 10, bottom: 10, left: 10, right: 10 } };

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`inline boxes retain native fallback, shared text, history, deletion, and saved geometry ${clients}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture(body),
        clientIds: clients,
      });
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!;
      let width = 40,
        revision = '1';
      const provider = (request: { tag: string | undefined }) =>
        request.tag === 'chip:source' ? { width, ascent: 20, descent: 4 } : null;
      const contexts = [0, 1].map(() => ({
        session: createLayoutSession(),
        cache: createParagraphLayoutCache<never>(),
      }));
      const layout = (part: OoxmlPart, peer = 0) => {
        const measurer: TextMeasurer = {
          ...createFixedMeasurer(),
          inlineWidgets: { measure: provider, revision },
        };
        return layoutSemanticDocument(part, 1, {
          ...contexts[peer]!,
          measurer,
          producer: 'fixed-fonts',
          geometry,
        });
      };
      const shape = (part: OoxmlPart, peer = 0) => {
        const current = layout(part, peer),
          widget = linesOf(current)
            .flatMap((line) => line.spans)
            .find((span) => span.inlineWidget)!;
        expect(widget).toBeDefined();
        expect(widget.text).toBe('source');
        return { layout: current, widget };
      };
      try {
        const id = a.session.paragraphIds()[0]!;
        const before = canonicalOoxmlFingerprint(a.session.part().root);
        expect(shape(a.session.part()).widget.box.width).toBe(30);
        expect(canonicalOoxmlFingerprint(a.session.part().root)).toBe(before);
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
          const current = shape(surface.session.part(), peer);
          expect(current.widget.box).toMatchObject({ x: 18, width: 30 });
          expect(current.widget.range).toMatchObject({ start: 3, end: 9 });
          expect(caretAt(current.layout, { paragraphId: id, offset: 9 })?.x).toBe(48);
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(id);
        expect(
          caretAt(shape(a.session.part()).layout, { paragraphId: id, offset: remote!.head.offset })
        ).not.toBeNull();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(paragraphTextOf(a.session.part(), id)).toBe('AsourceB!');
        expect(shape(a.session.part()).widget.box.x).toBe(6);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        width = 64;
        revision = '2';
        for (const [peer, surface] of [a, b].entries())
          expect(shape(surface.session.part(), peer).widget.box.width).toBe(48);
        pair.pause();
        const peerPoint = { paragraphId: id, offset: 11 };
        b.setSelection({ anchor: peerPoint, head: peerPoint });
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
        for (const [peer, editor] of pair.editors.entries()) {
          const part = editor.surface!.session.part();
          expect(paragraphTextOf(part, id)).toBe('aaAsourceB!?');
          expect(shape(part, peer).widget.box.width).toBe(48);
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          const reopened = saved.package.parts.get(saved.package.mainDocumentPart)!;
          const current = shape(reopened, peer);
          expect(current.widget.box).toMatchObject({ x: 18, width: 48 });
          expect(current.widget.range).toMatchObject({ start: 3, end: 9 });
        }
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
      } finally {
        pair.destroy();
      }
    });
