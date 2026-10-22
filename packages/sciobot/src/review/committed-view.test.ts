import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, collectRevisionSites } from '@docx-editor.dev/core/store';
import { reviewDocumentFixture } from '../../test/support/review-document';
import { markdownRangeAt } from '../projection/markdown';
import { readEditorRevision } from '../projection/revision-cache';
import { readCommittedView } from './committed-view';

test('committed reads hide insertions, restore deletions and original formatting, and preserve live state', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: reviewDocumentFixture() });
  try {
    const surface = editor.surface!;
    const session = surface.session;
    const revision = readEditorRevision(editor, 'doc', 'load-1')!;
    const before = canonicalOoxmlFingerprint(session.part());
    const sites = collectRevisionSites(session.part()).length;
    const history = [session.canUndo(), session.canRedo()];
    const result = await readCommittedView(session, revision, {
      measurer: surface.textMeasurer?.(),
    });
    expect(result.success).toBe(true);
    if (!result.success) throw new Error(result.error);
    expect(result.view.kind).toBe('committed');
    expect(result.view.projection.text).toBe('*Kept* Old\n\n## Heading');
    expect(result.view.layout.displayMode).toBe('original');
    expect(markdownRangeAt(result.view.projection, 7, 10)?.start).toEqual({
      sourceScope: 'body',
      paragraphId: session.paragraphIds()[0]!,
      offset: 5,
    });
    expect(result.view.revision).toEqual(revision);
    expect(result.view.revision === revision).toBe(false);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
    expect(session.packageRevision()).toBe(revision.revision);
    expect(collectRevisionSites(session.part()).length).toBe(sites);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('committed reads add no undo units and reflect later user edits', async () => {
  const container = document.createElement('div');
  const editor = createDocxEditor({ container, document: reviewDocumentFixture() });
  try {
    const surface = editor.surface!;
    const id = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    surface.insertPlainText('User: ');
    const read = () =>
      readCommittedView(surface.session, readEditorRevision(editor, 'doc', 'load-1')!);
    const typedFingerprint = canonicalOoxmlFingerprint(surface.session.part());
    const first = await read();
    const second = await read();
    expect(first.success && first.view.projection.text.includes('User: ')).toBe(true);
    expect(first).toEqual(second);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(typedFingerprint);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const undone = await read();
    expect(undone.success && undone.view.projection.text).toBe('*Kept* Old\n\n## Heading');
    expect(surface.session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const redone = await read();
    expect(redone.success && redone.view.projection.text.includes('User: ')).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(typedFingerprint);
  } finally {
    editor.destroy();
  }
});

test('stale, racing, and aborted reads return failures without resolving native proposals', async () => {
  const container = document.createElement('div');
  const editor = createDocxEditor({ container, document: reviewDocumentFixture() });
  try {
    const surface = editor.surface!;
    const token = readEditorRevision(editor, 'doc', 'load-1')!;
    const before = canonicalOoxmlFingerprint(surface.session.part());
    expect(
      await readCommittedView(surface.session, { ...token, revision: token.revision + 1 })
    ).toMatchObject({ success: false, code: 'stale-revision' });
    const controller = new AbortController();
    controller.abort();
    expect(
      await readCommittedView(surface.session, token, { signal: controller.signal })
    ).toMatchObject({ success: false, code: 'document-not-ready' });
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
    const racing = readCommittedView(surface.session, token);
    const id = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    surface.insertPlainText('Late: ');
    const landed = canonicalOoxmlFingerprint(surface.session.part());
    expect(await racing).toMatchObject({
      success: false,
      code: 'stale-revision',
      actualRevision: { revision: surface.session.packageRevision() },
    });
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(landed);
  } finally {
    editor.destroy();
  }
});
