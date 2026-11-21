/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { openTreeSession } from '@docx-editor.dev/core/binding';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support.ts';

test('automatic fit preserves each participant’s own viewport and document selection', async () => {
  const room = createPeerHarness('fit-scroll');
  const pair = await room.pair(zipDocument('<w:p><w:r><w:t>hello</w:t></w:r></w:p>'));
  function mount(peer: Peer) {
    peer.detach();
    const scroller = document.createElement('div');
    scroller.className = 'docx-editor__scroll-container';
    let heightReads = 0;
    Object.defineProperties(scroller, {
      clientWidth: { value: 1000 },
      clientHeight: {
        get() {
          heightReads++;
          return 600;
        },
      },
      scrollHeight: { value: 4000 },
    });
    const container = document.createElement('div');
    scroller.append(container);
    document.body.append(scroller);
    const editor = createDocxEditor({
      container,
      document: peer.room.document,
      modules: [collaborationModule({ session: peer.room.session })],
      zoom: 1,
    });
    return { editor, scroller, heightReads: () => heightReads };
  }
  const alice = mount(pair.alice),
    bob = mount(pair.bob);
  try {
    const id = room.paragraphIdAt(pair.alice, 0);
    const position = { paragraphId: id, offset: 2 };
    const range = { anchor: position, head: position };
    expect(alice.editor.exec({ type: 'setSelection', range }).ok).toBe(true);
    const nextFrame = () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    await nextFrame();
    const readsBeforeScroll = bob.heightReads();
    bob.scroller.scrollTop = 320;
    bob.scroller.dispatchEvent(new Event('scroll'));
    await nextFrame();
    expect(bob.heightReads()).toBe(readsBeforeScroll);
    const revision = alice.editor.surface!.session.packageRevision();
    expect(alice.editor.setZoomMode({ type: 'fit', fit: 'pageWidth' }).ok).toBe(true);
    expect(alice.editor.getZoom()).toBeGreaterThan(1);
    expect(alice.scroller.scrollTop).toBe(0);
    expect(alice.editor.surface!.state().selection).toMatchObject(range);
    expect(alice.editor.surface!.session.packageRevision()).toBe(revision);
    expect(bob.editor.getZoom()).toBe(1);
    expect(bob.scroller.scrollTop).toBe(320);
    expect(alice.editor.exec({ type: 'insertText', text: 'X' }).ok).toBe(true);
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(bob.editor.surface!.session.bodyText()).toBe('heXllo');
    const reopened = openTreeSession(new Uint8Array(await bob.editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(reopened.session.bodyText()).toBe('heXllo');
    expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(bob.editor.surface!.session.bodyText()).toBe('hello');
    expect(bob.scroller.scrollTop).toBe(320);
  } finally {
    for (const { editor, scroller } of [alice, bob]) {
      editor.destroy();
      scroller.remove();
    }
    room.cleanup();
  }
});
