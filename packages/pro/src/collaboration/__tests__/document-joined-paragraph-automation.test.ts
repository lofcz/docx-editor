/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { DocxEditor } from '@docx-editor.dev/editor-api/browser';
import { collaborationModule, reviewModule } from '../../index.ts';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const writerIndex of [0, 1])
    test(`joined Word paragraph replacement keeps both peer histories and native reopening: ${clients}/${writerIndex}`, async () => {
      const harness = createPeerHarness(
        'joined-paragraph-automation',
        { offlineEditing: true },
        { alice: clients[0], bob: clients[1] }
      );
      const pair = await harness.pair(
        zipDocument(
          '<w:p><w:pPr><w:rPr><w:del w:id="9" w:author="Other" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Hello world</w:t></w:r></w:p><w:p><w:r><w:t>Next</w:t></w:r></w:p>'
        )
      );
      const peers = [pair.alice, pair.bob].map((peer, index) => {
        peer.detach();
        const container = document.createElement('div');
        document.body.appendChild(container);
        const editor = createDocxEditor({
          container,
          document: peer.room.document,
          modules: [reviewModule(), collaborationModule({ session: peer.room.session })],
        });
        return {
          ...peer,
          container,
          editor,
          runtime: DocxEditor.createBrowser(editor, {
            author: index === 0 ? 'Author' : 'Reviewer',
          }),
        };
      });
      const flush = () => {
        for (const peer of peers) peer.room.session.flushPendingJournals();
      };
      const xml = async (index: number) =>
        strFromU8(
          unzipSync(new Uint8Array(await peers[index]!.editor.save()))['word/document.xml']!
        );
      const read = async (index: number) =>
        peers[index]!.runtime.run(async (c) => {
          const paragraphs = c.document.body.paragraphs;
          paragraphs.load('items/text');
          await c.sync();
          return paragraphs.items.map((paragraph) => paragraph.text);
        });
      try {
        await peers[0]!.runtime.run(async (c) => {
          c.document.changeTrackingMode = 'TrackMineOnly';
          c.document.body.search('world').getFirst().insertText('x\ny', 'Before');
          await c.sync();
        });
        flush();
        await peers[1]!.runtime.run(async (c) => {
          c.document.body.revisions.acceptAll();
          await c.sync();
        });
        flush();
        expect(await read(0)).toEqual(['Hello x', 'yworldNext']);
        expect(await read(1)).toEqual(['Hello x', 'yworldNext']);
        const baseline = await xml(0);
        expect(await xml(1)).toBe(baseline);
        // Keep this edit separate from the preceding actor's completed review decision.
        await new Promise((resolve) => setTimeout(resolve, 550));
        await peers[writerIndex]!.runtime.run(async (c) => {
          c.document.changeTrackingMode = 'Off';
          const range = c.document.body.paragraphs.getLast().search('worldNext').getFirst();
          range.insertText('changed', 'Replace');
          await c.sync();
        });
        flush();
        expect(await read(0)).toEqual(['Hello x', 'ychanged']);
        expect(await read(1)).toEqual(['Hello x', 'ychanged']);
        const changed = await xml(0);
        expect(await xml(1)).toBe(changed);
        expect(peers[writerIndex]!.editor.exec({ type: 'undo' }).ok).toBe(true);
        flush();
        for (const index of [0, 1]) {
          expect(await read(index)).toEqual(['Hello x', 'yworldNext']);
          expect(await xml(index)).toBe(baseline);
        }
        expect(peers[writerIndex]!.editor.exec({ type: 'redo' }).ok).toBe(true);
        flush();
        for (const index of [0, 1]) {
          expect(await xml(index)).toBe(changed);
          const bytes = new Uint8Array(await peers[index]!.editor.save());
          const reopened = await DocxEditor.createServer(bytes);
          try {
            await reopened.run(async (c) => {
              c.document.body.load('text');
              await c.sync();
              expect(c.document.body.text).toBe('Hello x\rychanged');
            });
            expect(strFromU8(unzipSync(await reopened.save())['word/document.xml']!)).toBe(changed);
          } finally {
            reopened.dispose();
          }
        }
      } finally {
        for (const peer of peers) {
          peer.runtime.dispose();
          peer.editor.destroy();
          peer.container.remove();
        }
        harness.cleanup();
      }
    });
