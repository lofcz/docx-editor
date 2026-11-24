/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { caretAt } from '@docx-editor.dev/core/layout';
import { mountPaginatedSurface } from '../../../../core/src/editor/paginated-surface.ts';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support.ts';

function mount(bytes: Uint8Array, peer?: Peer) {
  const scroll = document.createElement('div');
  scroll.className = 'docx-editor__scroll-container';
  Object.defineProperty(scroll, 'clientHeight', { value: 240 });
  const host = document.createElement('div');
  scroll.append(host);
  document.body.append(scroll);
  peer?.detach();
  const open = mountPaginatedSurface(host, bytes, {
    scale: 1,
    ...(peer && { collaborationModel: { session: peer.room.session } }),
  });
  if (!open.ok) throw Error(open.reason);
  const surface = open.surface;
  return {
    scroll,
    host,
    surface,
    close() {
      surface.destroy();
      scroll.remove();
    },
  };
}
test('scrolled collaborators retain editable pages, selections and saved edits', async () => {
  const harness = createPeerHarness('viewport-pair', { offlineEditing: true });
  const body = Array.from(
    { length: 12 },
    (_, i) =>
      '<w:p>' +
      (i ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : '') +
      '<w:r><w:t>Sheet ' +
      i +
      '</w:t></w:r></w:p>'
  ).join('');
  const peers = await harness.pair(zipDocument(body));
  const left = mount(peers.alice.room.document, peers.alice),
    right = mount(peers.bob.room.document, peers.bob);
  const xml = (view: typeof left) =>
    strFromU8(unzipSync(view.surface.save())['word/document.xml']!);
  const check = () => {
    peers.alice.room.session.flushPendingJournals();
    peers.bob.room.session.flushPendingJournals();
    expect(xml(left)).toBe(xml(right));
  };
  const select = (v: typeof left, page: number, start: number, end = start) => {
    const paragraphId = v.surface.session.paragraphIds()[page]!;
    v.surface.setSelection({
      anchor: { paragraphId, offset: start },
      head: { paragraphId, offset: end },
    });
  };
  try {
    for (const v of [left, right]) {
      expect(v.surface.layout().pages.length).toBe(12);
      v.scroll.scrollTop = v.surface.layout().pages[9]!.box.y + 5;
      v.scroll.dispatchEvent(new Event('scroll'));
      await new Promise((resolve) => setTimeout(resolve, 60));
      expect(v.host.querySelector('[data-page-index="9"] .docx-line')).not.toBeNull();
      expect(v.host.querySelectorAll('[data-materialized="true"]').length).toBeLessThan(12);
    }
    select(left, 9, 0, 3);
    applyAwarenessUpdate(
      peers.bob.awareness,
      encodeAwarenessUpdate(peers.alice.awareness, [peers.alice.ydoc.clientID]),
      'relay'
    );
    const remote = peers.bob.room.session.remoteSelections()[0]!;
    for (const key of ['anchor', 'head'] as const) {
      const point = { paragraphId: remote[key].nodeId, offset: remote[key].offset };
      const geometry = caretAt(right.surface.layout(), point);
      expect(geometry?.pageIndex).toBe(9);
      expect(geometry).toEqual(caretAt(left.surface.layout(), left.surface.state().selection[key]));
    }
    peers.pause();
    select(left, 9, 0, 1);
    left.surface.deleteForward();
    select(right, 9, 7);
    right.surface.type('!');
    peers.resume();
    check();
    expect(right.surface.session.bodyText()).toContain('heet 9!');
    left.surface.undo();
    check();
    expect(right.surface.session.bodyText()).toContain('Sheet 9!');
    left.surface.redo();
    check();
    expect(right.surface.session.bodyText()).toContain('heet 9!');
    const reopened = mount(right.surface.save());
    try {
      expect(xml(reopened)).toBe(xml(left));
    } finally {
      reopened.close();
    }
  } finally {
    left.close();
    right.close();
    harness.cleanup();
  }
});
