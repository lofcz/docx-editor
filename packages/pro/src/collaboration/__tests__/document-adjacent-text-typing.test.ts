/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { afterAll, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { collaborationModule } from '../../index.ts';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';

const registered = !GlobalRegistrator.isRegistered;
if (registered) GlobalRegistrator.register();
afterAll(() => {
  if (registered) GlobalRegistrator.unregister();
});

function text(editor: DocxEditorInstance) {
  const session = editor.surface!.session;
  return paragraphTextOf(session.part(), session.paragraphIds()[0]!);
}

function typeAt(editor: DocxEditorInstance, offset: number, value: string) {
  const surface = editor.surface!;
  const point = { paragraphId: surface.session.paragraphIds()[0]!, offset };
  surface.setSelection({ anchor: point, head: point });
  expect(editor.exec({ type: 'insertText', text: value }).ok).toBe(true);
  surface.flushPendingInput();
}

for (const [aliceId, bobId] of [
  [73, 41],
  [41, 73],
] as const) {
  test(`offline typing retains adjacent text once with client IDs ${aliceId}/${bobId}`, async () => {
    const harness = createPeerHarness('adjacent-text-typing', undefined, {
      alice: aliceId,
      bob: bobId,
    });
    const editors: DocxEditorInstance[] = [];
    try {
      const bytes = zipDocument(
        '<w:p><w:r><w:rPr><w:b/></w:rPr>' +
          '<w:t xml:space="preserve">First word </w:t><w:t>second word.</w:t></w:r></w:p>'
      );
      const { alice, bob, pause, resume } = await harness.pair(bytes);
      const [a, b] = [alice, bob].map((peer) => {
        peer.detach();
        const editor = createDocxEditor({
          container: document.createElement('div'),
          document: peer.room.document,
          modules: [collaborationModule({ session: peer.room.session })],
        });
        editors.push(editor);
        return editor;
      }) as [DocxEditorInstance, DocxEditorInstance];
      const flush = () => {
        alice.room.session.flushPendingJournals();
        bob.room.session.flushPendingJournals();
      };
      const expectBoth = (expected: string) => {
        flush();
        expect(text(a)).toBe(expected);
        expect(text(b)).toBe(expected);
      };
      pause();
      typeAt(a, 2, 'AA');
      typeAt(b, 5, 'BB');
      flush();
      expect(text(a)).toBe('FiAArst word second word.');
      expect(text(b)).toBe('FirstBB word second word.');
      resume();
      expectBoth('FiAArstBB word second word.');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expectBoth('FirstBB word second word.');
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      expectBoth('FiAArstBB word second word.');
      pause();
      const paragraphId = a.surface!.session.paragraphIds()[0]!;
      a.surface!.setSelection({
        anchor: { paragraphId, offset: 2 },
        head: { paragraphId, offset: 4 },
      });
      expect(a.exec({ type: 'deleteText' }).ok).toBe(true);
      typeAt(b, 'FiAArstBB word second word.'.length, '!');
      flush();
      expect(text(a)).toBe('FirstBB word second word.');
      expect(text(b)).toBe('FiAArstBB word second word.!');
      resume();
      expectBoth('FirstBB word second word.!');
      for (const peer of [a, b]) {
        const reopened = createDocxEditor({
          container: document.createElement('div'),
          document: new Uint8Array(await peer.save()),
        });
        editors.push(reopened);
        expect(text(reopened)).toBe('FirstBB word second word.!');
      }
    } finally {
      for (const editor of editors) editor.destroy();
      harness.cleanup();
    }
  });
}
