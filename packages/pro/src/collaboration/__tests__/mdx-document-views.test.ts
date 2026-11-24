/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { afterEach, expect, test } from 'bun:test';
import { applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { createDocxEditor, setSurfaceInlineWidgetMetrics } from '@docx-editor.dev/core/editor';
import { contentControlTextOf } from '@docx-editor.dev/core/store';
import { caretAt, hitTestSemantic } from '@docx-editor.dev/core/layout';
import {
  configureMdxDisclosures,
  exportMdxDocument,
  importMdxDocument,
  mountMdxComponents,
  mountMdxDisclosures,
  readMdxComponents,
} from '../../../../docx-to-markdown/src/mdx.ts';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, type Peer } from './document-peer-support.ts';
import { loadPackage, packageDigest, packageFingerprint } from './document-support.ts';

const source =
  '<details open>\n<summary>\n\n| Limits | Unit |\n| --- | --- |\n| Maximum | kPa |\n\nSummary detail\n\n</summary>\n\n<ProjectPlot>\n\nObserved value\n\n</ProjectPlot>\n\n</details>\n\nShared tail';
const harness = createPeerHarness('mdx-document-views', { offlineEditing: true });
const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const close of cleanups.splice(0).reverse()) close();
  harness.cleanup();
});

function mount(bytes: Uint8Array, peer?: Peer) {
  peer?.detach();
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: bytes,
    configureInitialBlockView: configureMdxDisclosures,
    zoom: 1,
    modules: peer ? [collaborationModule({ session: peer.room.session })] : [],
  });
  cleanups.push(() => {
    editor.destroy();
    container.remove();
  });
  const surface = editor.surface!;
  cleanups.push(mountMdxDisclosures(editor));
  cleanups.push(
    mountMdxComponents(
      editor,
      new Map([
        [
          'ProjectPlot',
          ({ element, observeContent }) => {
            const button = element.ownerDocument.createElement('button');
            button.textContent = 'Inspect plot';
            element.append(button);
            return observeContent((control) => {
              button.title = contentControlTextOf(control);
            });
          },
        ],
      ])
    )
  );
  surface.flushPendingInput();
  const fold = () => container.querySelector<HTMLButtonElement>('button[aria-label="Limits"]')!;
  return { editor, surface, container, fold, pkg: () => surface.session.currentPackage() };
}

type Mounted = ReturnType<typeof mount>;
async function room(mdx = source) {
  const pair = await harness.pair(importMdxDocument(mdx));
  const alice = mount(pair.alice.room.document, pair.alice);
  const bob = mount(pair.bob.room.document, pair.bob);
  const flush = () => {
    for (const peer of [pair.alice, pair.bob]) peer.room.session.flushPendingJournals();
    for (const host of [alice, bob]) host.surface.flushPendingInput();
  };
  const converge = () => {
    flush();
    expect(packageFingerprint(alice.pkg())).toBe(packageFingerprint(bob.pkg()));
    expect(exportMdxDocument(alice.pkg())).toBe(exportMdxDocument(bob.pkg()));
  };
  return { pair, alice, bob, flush, converge };
}

function point(host: Mounted, text: string, offset: number) {
  const [match] = host.editor.findMatches(text);
  if (!match) throw new Error('Missing paragraph containing ' + text);
  return { paragraphId: match.blockId, offset: match.start + offset };
}

function visiblePoint(host: Mounted, text: string, offset: number) {
  const layout = host.surface.layout();
  const position = point(host, text, offset);
  const caret = caretAt(layout, position);
  if (!caret) throw new Error('Missing visible caret for ' + text);
  const hit = hitTestSemantic(layout, {
    pageIndex: caret.pageIndex,
    x: caret.x,
    y: caret.y + caret.height / 2,
  });
  expect(hit?.position).toEqual(position);
  return hit!.position;
}

function write(host: Mounted, text: string, offset: number, inserted: string) {
  const position = visiblePoint(host, text, offset);
  host.surface.setSelection({ anchor: position, head: position });
  expect(host.editor.exec({ type: 'insertText', text: inserted }).ok).toBe(true);
}

test('unchanged nested views and disclosure anchors stay mounted during peer edits', async () => {
  const blocks = Array.from(
    { length: 8 },
    (_, i) =>
      '<details open>\n<summary>Panel ' +
      i +
      '</summary>\n\n' +
      '<UnknownGroup>\n\n<ProjectPlot>\n\nPlot ' +
      i +
      '\n\n</ProjectPlot>\n\n' +
      '<details open>\n<summary>Nested ' +
      i +
      '</summary>\n\nShared body\n\n</details>\n\n</UnknownGroup>\n\n</details>'
  );
  const { alice, bob, flush, converge } = await room([...blocks, 'Tail text'].join('\n\n'));
  const hosts = [alice, bob];
  const views = hosts.map(({ container }) =>
    container.querySelector('[data-mdx-component="ProjectPlot"]')
  );
  const anchors = hosts.map(({ container }) =>
    container.querySelector('button[aria-label="Nested 0"]')
  );
  expect(views.every(Boolean)).toBe(true);
  expect(anchors.every(Boolean)).toBe(true);
  const observers = views.map((view) => {
    const observer = new MutationObserver(() => {});
    observer.observe(view!.parentElement!, { childList: true });
    return observer;
  });
  const layers = hosts.map(({ container }) => container.querySelector('.docx-pages')!);
  const scans = { pageBodies: 0, summaries: 0 };
  const original = Element.prototype.querySelectorAll;
  try {
    Element.prototype.querySelectorAll = function (selector: string) {
      if (layers.includes(this)) {
        if (selector === '.docx-page-content') scans.pageBodies++;
        if (selector === '.docx-paragraph-fragment[data-fragment-index="0"]') scans.summaries++;
      }
      return original.call(this, selector);
    };
    for (let i = 0; i < 10; i++) {
      write(hosts[i % 2]!, 'Tail text', 0, String(i));
      flush();
    }
  } finally {
    Element.prototype.querySelectorAll = original;
    for (const [i, observer] of observers.entries()) {
      const removed = observer.takeRecords().flatMap((event) => [...event.removedNodes]);
      observer.disconnect();
      expect(removed).not.toContain(views[i]);
    }
  }
  converge();
  write(alice, 'Plot 0', 0, 'Nested ');
  flush();
  write(bob, 'Nested 0', 0, 'Changed ');
  converge();
  for (const [i, host] of hosts.entries()) {
    expect(host.container.querySelector('[data-mdx-component="ProjectPlot"]')).toBe(views[i]);
    expect(views[i]?.querySelector('button')?.title).toBe('Nested Plot 0');
    expect(host.container.querySelector('button[aria-label="Changed Nested 0"]')).toBe(anchors[i]);
    expect(exportMdxDocument(host.pkg())).toContain('0123456789Tail text');
  }
  expect(scans).toEqual({ pageBodies: 0, summaries: 0 });
});

async function reopen(host: Mounted) {
  const saved = new Uint8Array(await host.editor.save());
  expect(packageDigest(loadPackage(saved))).toEqual(packageDigest(host.pkg()));
  const reopened = mount(saved);
  expect(exportMdxDocument(reopened.pkg())).toBe(exportMdxDocument(host.pkg()));
  return reopened;
}

test('each editor uses its own folds and visible selection geometry during concurrent typing and history', async () => {
  const { pair, alice, bob, flush, converge } = await room(
    source.replace('</details>', 'Unit <Term>symbol</Term>\n\n</details>')
  );
  const original = exportMdxDocument(alice.pkg());
  const fingerprint = packageFingerprint(alice.pkg());
  // Both layouts have been visited when the next shared edit arrives.
  alice.fold().click();
  flush();
  alice.fold().click();
  flush();
  const symbol = readMdxComponents(alice.pkg()).find(({ name }) => name === 'Term')!;
  setSurfaceInlineWidgetMetrics(
    alice.surface,
    new Map([[symbol.controlId, { width: 64, ascent: 48, descent: 4 }]])
  );
  alice.fold().click();
  flush();
  alice.fold().click();
  flush();
  expect(alice.surface.state().perf.placed).toBe(0);
  alice.fold().click();
  converge();
  expect(packageFingerprint(alice.pkg())).toBe(fingerprint);
  expect(alice.fold().getAttribute('aria-expanded')).toBe('false');
  expect(bob.fold().getAttribute('aria-expanded')).toBe('true');
  write(alice, 'kPa', 0, '1 ');
  converge();
  expect(visiblePoint(bob, '1 kPa', 2).offset).toBe(2);
  expect(alice.fold().getAttribute('aria-expanded')).toBe('false');
  expect(bob.fold().getAttribute('aria-expanded')).toBe('true');
  expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(exportMdxDocument(bob.pkg())).toBe(original);
  const plot = '[data-mdx-component="ProjectPlot"]';
  expect(alice.container.querySelector(plot)).toBeNull();
  expect(bob.container.querySelector(plot)).not.toBeNull();
  const button = bob.container.querySelector<HTMLButtonElement>(plot + ' button')!;
  expect(button.title).toBe('Observed value');
  expect(visiblePoint(alice, 'Summary detail', 2).offset).toBe(2);
  expect(caretAt(alice.surface.layout(), point(alice, 'Observed value', 0))).toBeNull();
  expect(caretAt(bob.surface.layout(), point(bob, 'Observed value', 0))).not.toBeNull();
  const head = visiblePoint(alice, 'Shared tail', 6);
  const anchor = visiblePoint(alice, 'Shared tail', 0);
  alice.surface.setSelection({ anchor, head });
  expect(alice.editor.query({ type: 'selectedText' })).toBe('Shared');
  const publishPresence = () =>
    applyAwarenessUpdate(
      pair.bob.awareness,
      encodeAwarenessUpdate(pair.alice.awareness, [pair.alice.ydoc.clientID]),
      'relay'
    );
  publishPresence();
  const [remote] = pair.bob.room.session.remoteSelections();
  expect(remote).toMatchObject({
    anchor: { nodeId: anchor.paragraphId, offset: anchor.offset },
    head: { nodeId: head.paragraphId, offset: head.offset },
  });
  const receiver = caretAt(bob.surface.layout(), {
    paragraphId: remote!.head.nodeId,
    offset: remote!.head.offset,
  })!;
  expect(receiver.y).toBeGreaterThan(caretAt(alice.surface.layout(), head)!.y);
  const band = bob.container.querySelector<HTMLElement>('.docx-remote-selection-rect')!;
  expect(parseFloat(band.style.width)).toBeGreaterThan(0);
  alice.surface.setSelection({ anchor: head, head });
  publishPresence();
  const remoteCaret = bob.container.querySelector<HTMLElement>('.docx-remote-caret')!;
  const page = bob.surface.layout().pages[receiver.pageIndex]!;
  // Both editors use 100% zoom. The receiver paints its own page coordinates in CSS pixels.
  expect(parseFloat(remoteCaret.style.top)).toBeCloseTo(
    ((page.contentBox.y + receiver.y) * 96) / 72,
    4
  );

  pair.pause();
  write(alice, 'Shared tail', 0, 'Alice ');
  write(bob, 'Observed value', 0, 'Bob ');
  flush();
  expect(alice.editor.findMatches('Bob Observed')).toHaveLength(0);
  expect(bob.editor.findMatches('Alice Shared')).toHaveLength(0);
  pair.resume();
  converge();
  const both = original
    .replace('Observed value', 'Bob Observed value')
    .replace('Shared tail', 'Alice Shared tail');
  expect(exportMdxDocument(alice.pkg())).toBe(both);
  expect(bob.container.querySelector(plot + ' button')).toBe(button);
  expect(button.title).toBe('Bob Observed value');
  expect(alice.fold().getAttribute('aria-expanded')).toBe('false');
  expect(bob.fold().getAttribute('aria-expanded')).toBe('true');
  expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(exportMdxDocument(bob.pkg())).toBe(
    original.replace('Observed value', 'Bob Observed value')
  );
  expect(alice.editor.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(exportMdxDocument(bob.pkg())).toBe(both);
  for (const host of [alice, bob]) {
    const saved = await reopen(host);
    expect(saved.fold().getAttribute('aria-expanded')).toBe('true');
    expect(visiblePoint(saved, 'Bob Observed value', 0).offset).toBe(0);
  }
});

test('removing a project control while its child is edited removes the edit and both trusted views', async () => {
  const { pair, alice, bob, flush, converge } = await room();
  const plot = readMdxComponents(alice.pkg()).find(({ name }) => name === 'ProjectPlot')!;
  const view = '[data-mdx-component="ProjectPlot"]';
  const originals = [alice, bob].map((host) => host.container.querySelector(view)!);
  expect(originals.every((el) => el?.isConnected)).toBe(true);
  pair.pause();
  expect(
    alice.surface.session.applyTreeOps([
      { op: 'removeContentControl', controlId: plot.controlId, keepContent: false },
    ]).committed
  ).toBe(true);
  write(bob, 'Observed value', 0, 'Concurrent ');
  flush();
  expect(bob.editor.findMatches('Concurrent Observed value')).toHaveLength(1);
  expect(alice.container.querySelector(view)).toBeNull();
  pair.resume();
  converge();
  for (const host of [alice, bob]) {
    expect(readMdxComponents(host.pkg()).map(({ name }) => name)).toEqual(['details', 'summary']);
    expect(host.editor.findMatches('Observed value')).toHaveLength(0);
    expect(host.editor.findMatches('Concurrent')).toHaveLength(0);
    expect(host.container.querySelector(view)).toBeNull();
    const saved = await reopen(host);
    expect(saved.editor.findMatches('Shared tail')).toHaveLength(1);
    expect(readMdxComponents(saved.pkg()).map(({ name }) => name)).toEqual(['details', 'summary']);
  }
  expect(originals.every((el) => !el.isConnected)).toBe(true);
});
