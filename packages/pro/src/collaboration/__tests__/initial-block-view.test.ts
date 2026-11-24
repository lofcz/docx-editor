/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import { createDocxEditor, type ConfigureInitialBlockView } from '@docx-editor.dev/core/editor';
import {
  importMdxDocument,
  readMdxComponents,
  exportMdxDocument,
} from '../../../../docx-to-markdown/src/mdx.ts';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, type Peer } from './document-peer-support.ts';
import { loadPackage, packageDigest } from './document-support.ts';

test('initial widget height, folds, and insets stay local through peer edits, history, and saving', async () => {
  const harness = createPeerHarness('initial-block-view');
  const { alice, bob } = await harness.pair(
    importMdxDocument('<details>\n<summary>Summary</summary>\n\nAnswer\n\n</details>')
  );
  function mount(peer: Peer, configureInitialBlockView?: ConfigureInitialBlockView) {
    peer.detach();
    const container = document.createElement('div');
    const editor = createDocxEditor({
      container,
      document: peer.room.document,
      modules: [collaborationModule({ session: peer.room.session })],
      ...(configureInitialBlockView ? { configureInitialBlockView } : {}),
    });
    return { editor, surface: editor.surface!, container };
  }
  const folded = mount(alice, (pkg, view) => {
    const detail = readMdxComponents(pkg).find((v) => v.name === 'details')!;
    view.setInitialBlockHeights(new Map([[detail.summaryParagraphId!, (width) => width / 4]]));
    expect(
      view.setCollapsedContentControls(new Map([[detail.controlId, detail.summaryParagraphId!]]))
    ).toEqual({ ok: true });
    view.setBlockInsets(new Map([[detail.summaryParagraphId!, { left: 12, right: 0 }]]));
  });
  const open = mount(bob);
  try {
    const flush = () => {
      alice.room.session.flushPendingJournals();
      bob.room.session.flushPendingJournals();
      folded.surface.flushPendingInput();
      open.surface.flushPendingInput();
    };
    const check = (answer: string) => {
      flush();
      expect(folded.container.textContent).not.toContain('Answer');
      expect(open.container.textContent).toContain(answer);
      expect(folded.surface.collapsedContentControls()).toHaveLength(1);
      expect(open.surface.collapsedContentControls()).toEqual([]);
      expect(exportMdxDocument(folded.surface.session.currentPackage())).toContain(answer);
    };
    check('Answer');
    const [match] = open.editor.findMatches('Answer');
    const point = { paragraphId: match!.blockId, offset: match!.start + 'Answer'.length };
    open.surface.setSelection({ anchor: point, head: point });
    expect(open.editor.exec({ type: 'insertText', text: ' edited' }).ok).toBe(true);
    check('Answer edited');
    expect(open.editor.exec({ type: 'undo' }).ok).toBe(true);
    check('Answer');
    expect(open.editor.exec({ type: 'redo' }).ok).toBe(true);
    check('Answer edited');
    const saved = loadPackage(new Uint8Array(await folded.editor.save()));
    expect(exportMdxDocument(saved)).toContain('Answer edited');
    expect(packageDigest(saved)).toEqual(
      packageDigest(loadPackage(new Uint8Array(await open.editor.save())))
    );
  } finally {
    folded.editor.destroy();
    open.editor.destroy();
    harness.cleanup();
  }
});
